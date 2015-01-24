using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class ForBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public ForBlockTranslator(
            CSharpName supportRefName,
            CSharpName envClassName,
            CSharpName envRefName,
            CSharpName outerClassName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
			ITranslateIndividualStatements statementTranslator,
			ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
            : base(supportRefName, envClassName, envRefName, outerClassName, outerRefName, nameRewriter, tempNameGenerator, statementTranslator, valueSettingStatementTranslator, logger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
            if (logger == null)
                throw new ArgumentNullException("logger");

            _statementTranslator = statementTranslator;
            _logger = logger;
        }

		public TranslationResult Translate(ForBlock forBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (forBlock == null)
                throw new ArgumentNullException("forBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // A note about ON ERROR RESUME NEXT and the intriguing behaviour it can introduce. The following will display "We're in the loop! i is Empty" -
            //
            //   On Error Resume Next
            //   Dim i: For i = 1 To 1/0
            //     WScript.Echo "We're in the loop! i is " & TypeName(i)
            //   Next
            //
            // If error-trapping is enabled and one of the From, To or Step evaluation fails, no further constaints will be receive evaluation attempts but
            // the loop WILL be entered.. once. Only once. And the loop variable will not be initialised when doing so (so, in the above example, it has
            // the "Empty" value inside the loop - but if it had been set to "a" before the loop construct then it would still have value "a" within
            // the loop).

            // Identify tokens for the start, end and step variables. If they are numeric constants then use them, otherwise they must be stored in temporary
            // values. Note that these temporary values are NOT re-evaluated each loop since this is how VBScript (unlike some other languages) work.
            var undeclaredVariableReferencesAccessedByLoopConstraints = new NonNullImmutableList<NameToken>();
            var loopConstraintInitialisersWhereRequired = new List<Tuple<CSharpName, string>>();
            string loopStart;
            var numericLoopStartValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopFrom);
            if (numericLoopStartValueIfAny != null)
                loopStart = numericLoopStartValueIfAny.AsCSharpValue();
            else
            {
                // If the start value is not a simple numeric constant then we'll need to declare a variable for it. This variable will never need to be
                // accessed elsewhere so it doesn't need to be declared as a real VariableDeclaration (and none of the temporary variables - such as
                // function return values, for example - need to be added to the ScopeAccessInformation since they are guaranteed to be unique).
                var loopStartExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopFrom,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );
                var loopStartName = _tempNameGenerator(new CSharpName("loopStart"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(Tuple.Create(
                    loopStartName,
                    string.Format(
                        "{0}.CDBL({1})",
                        _supportRefName.Name,
                        loopStartExpressionContent.TranslatedContent
                    )
                ));
                loopStart = loopStartName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    loopStartExpressionContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }
            string loopEnd;
            var numericLoopEndValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopTo);
            if (numericLoopEndValueIfAny != null)
                loopEnd = numericLoopEndValueIfAny.AsCSharpValue();
            else
            {
                // Same logic as for the loopStart value above applies here
                var loopEndExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopTo,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );
                var loopEndName = _tempNameGenerator(new CSharpName("loopEnd"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(Tuple.Create(
                    loopEndName,
                    string.Format(
                        "{0}.CDBL({1})",
                        _supportRefName.Name,
                        loopEndExpressionContent.TranslatedContent
                    )
                ));
                loopEnd = loopEndName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    loopEndExpressionContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }
            string loopStep;
            var numericLoopStepValueIfAny = (forBlock.LoopStep == null)
                ? new NumericValueToken("1", forBlock.LoopTo.Tokens.Last().LineIndex) // Default to Step 1 if no LoopStep expression was specified
                : TryToGetExpressionAsNumericConstant(forBlock.LoopStep);
            if ((numericLoopStepValueIfAny == null) && (forBlock.LoopStep.Tokens.Count() == 2))
            {
                // If the loop step is "-1" then this will still appear as distinct tokens "-" and "1" since the NumberRebuilder can't safely combine
                // the two tokens at the level of knowledge that it has, since it doesn't know if the "STEP" token is a keyword (meaning it's part of
                // a loop structure and so they could be safely combined) or a reference name (meaning it could be something like "step - 1", in which
                // case they could not be safely combined).
                var firstLoopStepToken = forBlock.LoopStep.Tokens.First();
                var lastLoopStepTokenAsNumericValueToken = forBlock.LoopStep.Tokens.Last() as NumericValueToken;
                if ((firstLoopStepToken is OperatorToken) && (firstLoopStepToken.Content == "-") && (lastLoopStepTokenAsNumericValueToken != null))
                    numericLoopStepValueIfAny = lastLoopStepTokenAsNumericValueToken.GetNegative();
            }
            if (numericLoopStepValueIfAny != null)
                loopStep = numericLoopStepValueIfAny.AsCSharpValue();
            else
            {
                // Same logic as for the loopStart/loopEnd value above applies here
                var loopStepExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopStep,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );
                var loopStepName = _tempNameGenerator(new CSharpName("loopStep"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(Tuple.Create(
                    loopStepName,
                    string.Format(
                        "{0}.CDBL({1})",
                        _supportRefName.Name,
                        loopStepExpressionContent.TranslatedContent
                    )
                ));
                loopStep = loopStepName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    loopStepExpressionContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            // Any dynamic loop constraints (ie. those that can be confirmed to be fixed numeric values at translation time) need to have variables
            // declared and initialised. There may be a mix of dynamic and constant constraints so there may be zero, one, two or three variables
            // to deal with here (this will have been determined in the work above).
            // - Note: These fixed constraints are stored as as doubles (where error-trapping is involved, the variables are explicitly declared as
            //   double, though otherwise var is used when CDBL is called). This means that the loop termination conditions are simpler, since we
            //   only need to call NUM on the loop variable - which may be set to anything by code within the loop - we can rest safe in the knowledge
            //   that the constraints are always numeric since they are doubles. If the loop variable is set to a date type within the loop (which is
            //   valid VBScript), then the loop constraints comparisons will be fine (StrictLTE can compare a date to a double, for example). We can't
            //   use NUM to evaluate the contraints, since that returns an object (since it might return an int, a double, a date.. anything that
            //   VBScript considers to be a numeric type).
            var translationResult = TranslationResult.Empty;
            CSharpName constraintsInitialisedFlagNameIfAny;
            if (!loopConstraintInitialisersWhereRequired.Any())
                constraintsInitialisedFlagNameIfAny = null;
            else if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
            {
                constraintsInitialisedFlagNameIfAny = null;
                foreach (var loopConstraintInitialiser in loopConstraintInitialisersWhereRequired)
                {
                    translationResult = translationResult
                        .Add(new TranslatedStatement(
                            string.Format(
                                "var {0} = {1};",
                                loopConstraintInitialiser.Item1.Name,
                                loopConstraintInitialiser.Item2
                            ),
                            indentationDepth
                        ));
                }
            }
            else
            {
                constraintsInitialisedFlagNameIfAny = _tempNameGenerator(new CSharpName("loopConstraintsInitialised"), scopeAccessInformation);
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "double {0};", // See notes above as to why these are doubles (all dynamic loop constraints are evaluated to doubles)
                            string.Join(", ", loopConstraintInitialisersWhereRequired.Select(c => c.Item1.Name + " = 0"))
                        ),
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement(
                        string.Format(
                            "var {0} = false;",
                            constraintsInitialisedFlagNameIfAny.Name
                        ),
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement(
                        string.Format(
                            "{0}.HANDLEERROR({1}, () => {{",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ));
                foreach (var loopConstraintInitialiser in loopConstraintInitialisersWhereRequired)
                {
                    translationResult = translationResult
                        .Add(new TranslatedStatement(
                            string.Format(
                                "{0} = {1};",
                                loopConstraintInitialiser.Item1.Name,
                                loopConstraintInitialiser.Item2
                            ),
                            indentationDepth + 1
                        ));
                }
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        constraintsInitialisedFlagNameIfAny.Name + " = true;",
                        indentationDepth + 1
                    ))
                    .Add(new TranslatedStatement("});", indentationDepth));
            }

            // If all three constraints are numeric constraints then we can determine now whether the loop should be executed or not
            // - If the loop is descending but a negative step is not specified then the loop will not be entered and so might as well not be
            //   included in the translated content
            // - If the loop is ascending then a positive OR zero step must be specified
            var guardClauseLines = new NonNullImmutableList<string>();
            if ((numericLoopStartValueIfAny != null) && (numericLoopEndValueIfAny != null) && (numericLoopStepValueIfAny != null))
            {
                if ((numericLoopStartValueIfAny.Value > numericLoopEndValueIfAny.Value) && (numericLoopStepValueIfAny.Value >= 0))
                {
                    _logger.Warning("Optimising out descending FOR loop that does not have a negative step specified");
                    return TranslationResult.Empty;
                }
                if ((numericLoopStartValueIfAny.Value <= numericLoopEndValueIfAny.Value) && (numericLoopStepValueIfAny.Value < 0))
                {
                    _logger.Warning("Optimising out ascending FOR loop that has a negative step specified");
                    return TranslationResult.Empty;
                }
            }
            else if (numericLoopStepValueIfAny != null)
            {
                // If the step is a known numeric constant, then the guard clause only needs to compare the from and to constaints (at least
                // one of which must not be a known constant, otherwise we'd be in the above condition)
                // - Note: loopStart and loopEnd are both variables of type double, so we know we can do a straight comparison and don't have to
                //   rely upon methods such as StrictLTE
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    // Ascending loop or infinite loop (step zero, which is supported in VBScript), start must not be greater than end
                    guardClauseLines = guardClauseLines.Add(string.Format("({0} <= {1})", loopStart, loopEnd));
                }
                else
                {
                    // Descending loop, start must be greater than end
                    guardClauseLines = guardClauseLines.Add(string.Format("({0} > {1})", loopStart, loopEnd));
                }
            }
            else if ((numericLoopStartValueIfAny != null) && (numericLoopEndValueIfAny != null))
            {
                // If the from and to bounds are known to be constants, but not the step then we just need to ensure that the step matches the
                // direction of from and to at runtime. Note: A step of zero will cause an infinite loop, but only if from <= to (the loop will
                // not be executed if it is descending and has a step of zero)
                // - Note: loopStart and loopEnd are both variables of type double, so we know we can do a straight comparison and don't have to
                //   rely upon methods such as StrictLTE
                if (numericLoopStartValueIfAny.Value <= numericLoopEndValueIfAny.Value)
                    guardClauseLines = guardClauseLines.Add(string.Format("({0} >= 0)", loopStep));
                else
                    guardClauseLines = guardClauseLines.Add(string.Format("({0} < 0)", loopStep));
            }
            else
            {
                // There are no more shortcuts now, we need to check at runtime that loopStep is negative for a descending loop and non-negative
                // for a non-descending loop
                // - Note: loopStart and loopEnd are both variables of type double, so we know we can do a straight comparison and don't have to
                //   rely upon methods such as StrictLTE
                guardClauseLines = guardClauseLines.Add(string.Format(
                    "((({0} <= {1}) && ({2} >= 0))",
                    loopStart,
                    loopEnd,
                    loopStep
                ));
                guardClauseLines = guardClauseLines.Add(string.Format(
                    "|| (({0} > {1}) && ({2} < 0)))",
                    loopStart,
                    loopEnd,
                    loopStep
                ));
            }

            // If there are non-constant constraint(s) and error-trapping is enabled and thee constraint evaluation fails, then VBScript will
            // enter the loop once (only once, and it will not initialise the loop variable when doing so). In this case, we need to bypass
            // the guard clause - it won't make sense anyway to compare loop constaints if we know that the constraint evaluation failed!
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                // If the guard clases have already been split up over multiple lines then insert a fresh line, otherwise just bang it in
                // front of the single line (unless there is no content, in which case this will become the only guard clause line)
                if (!guardClauseLines.Any())
                    guardClauseLines = guardClauseLines.Add("(" + constraintsInitialisedFlagNameIfAny.Name + ")");
                else if (guardClauseLines.Count == 1)
                {
                    guardClauseLines = new NonNullImmutableList<string>(new[]
                    {
                        string.Format(
                            "(!{0} || {1})",
                            constraintsInitialisedFlagNameIfAny.Name,
                            guardClauseLines.Single()
                        )
                    });
                }
                else
                {
                    guardClauseLines = new NonNullImmutableList<string>(
                        new[] { "(!" + constraintsInitialisedFlagNameIfAny.Name + " ||" }
                        .Concat(guardClauseLines.Take(guardClauseLines.Count - 1))
                        .Concat(new[] { guardClauseLines.Last() + ")" })
                    );
                }
            }

            var rewrittenLoopVariableName = _nameRewriter.GetMemberAccessTokenName(forBlock.LoopVar);
            var targetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenLoopVariableName, _envRefName, _outerRefName, _nameRewriter);
            if (targetContainer != null)
                rewrittenLoopVariableName = targetContainer.Name + "." + rewrittenLoopVariableName;

            if (guardClauseLines.Any())
            {
                translationResult = translationResult.Add(new TranslatedStatement("if " + guardClauseLines.First(), indentationDepth));
                foreach (var guardClauseLine in guardClauseLines.Skip(1))
                    translationResult = translationResult.Add(new TranslatedStatement(guardClauseLine, indentationDepth));
                translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                indentationDepth++;
            }
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "{0}.HANDLEERROR({1}, () => {{",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ));
                indentationDepth++;
            }
            string continuationCondition;
            if (numericLoopStepValueIfAny != null)
            {
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    continuationCondition = string.Format(
                        "{0}.StrictLTE({1}, {2})",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        loopEnd
                    );
                }
                else
                {
                    continuationCondition = string.Format(
                        "{0}.StrictGTE({1}, {2})",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        loopEnd
                    );
                }
            }
            else
            {
                // Note: loopStep is a variable of type double, so we know we can do straight comparisons between it and zero, we don't need to rely
                // upon methods such as StrictLTE
                continuationCondition = string.Format(
                    "(({3} >= 0) && {0}.StrictLTE({1}, {2})) || (({3} < 0) && {0}.StrictGTE({1}, {2}))",
                    _supportRefName.Name,
                    rewrittenLoopVariableName,
                    loopEnd,
                    loopStep
                );
            }
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                // If there is a chance that the constraint evaluation will fail but that processing may continue (ie. there's an ON ERROR RESUME NEXT
                // in the current scope that may affect things) then the continuation condition gets a special condition that it will not be checked
                // if the constraint evaluation did indeed fail. In such a case, VBScript will enter the loop once (without altering the loop variable
                // value). To replicate this behaviour, if constraint evaluation fails then the loop WILL be entered once without the loop variable
                // being altered and without the continuation condition being considered - instead, there is a "break" at the end of the loop which
                // will be executed if the constraint evaluation failed, ensuring that the loop was indeed processed once, consistent with VBScript.
                continuationCondition = string.Format(
                    "!{0} || ({1})",
                    constraintsInitialisedFlagNameIfAny.Name,
                    continuationCondition
                );
            }
            string loopIncrementWithLeadingSpaceIfNonBlank;
            if ((numericLoopStepValueIfAny != null) && (numericLoopStepValueIfAny.Value == 0))
                loopIncrementWithLeadingSpaceIfNonBlank = "";
            else
            {
                if ((numericLoopStepValueIfAny != null) && (numericLoopStepValueIfAny.Value < 0))
                {
                    // If the step is known to be a negative numeric constant value, then use SUBT instead of ADD. If it's a small step value then it
                    // will cast to an Int16, so using SUBT instead of ADD allows the following (slight) improvement to readability:
                    //   _env.i = _.ADD(_env.i, (Int16)(-1))
                    //   _env.i = _.SUBT(_env.i, (Int16)1)
                    loopIncrementWithLeadingSpaceIfNonBlank = string.Format(
                        " {0} = {2}.SUBT({0}, {1})",
                        rewrittenLoopVariableName,
                        numericLoopStepValueIfAny.GetNegative().AsCSharpValue(),
                        _supportRefName.Name
                    );
                }
                else
                {
                    loopIncrementWithLeadingSpaceIfNonBlank = string.Format(
                        " {0} = {2}.ADD({0}, {1})",
                        rewrittenLoopVariableName,
                        loopStep,
                        _supportRefName.Name
                    );
                }
            }
            string loopVarInitialiser;
            if (constraintsInitialisedFlagNameIfAny == null)
            {
                loopVarInitialiser = string.Format(
                    "{0} = {1}",
                    rewrittenLoopVariableName,
                    loopStart
                );
            }
            else
            {
                // If error-trapping is enabled and loop constaint evaluation fails, then the loop must be entered once but the loop variable's value may
                // not be altered. This condition deals with that case; the loop variable is only set to the start value if all of the constraints were
                // successfully initialised. Note that constraintsInitialisedFlagNameIfAny will only be non-null if error-trapping may take affect and
                // if the loop constraints are not numeric constants known at translation time.
                loopVarInitialiser = string.Format(
                    "{0} = {1} ? {2} : {0}",
                    rewrittenLoopVariableName,
                    constraintsInitialisedFlagNameIfAny.Name,
                    loopStart
                );
            }
            if ((constraintsInitialisedFlagNameIfAny == null) && (numericLoopStepValueIfAny != null))
            {
                // If there is no complicated were-loop-constraints-successfully-initialised logic to worry about (meaning that either error-trapping was
                // not enabled or that all of the constraints were numeric constants), then render a nice and simple single-line loop construct.
                // Note: There is no space before {2} so that if there is no loop increment required then the output doesn't look like it's missing
                // something (this may be the case if the loop step is zero)
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(
                        "for ({0}; {1};{2})",
                        loopVarInitialiser,
                        continuationCondition,
                        loopIncrementWithLeadingSpaceIfNonBlank
                    ),
                    indentationDepth
                ));
            }
            else
            {
                // If there IS some possibility that loop constraint evaluation may fail at runtime, then the loop construct becomes much more complicated
                // and so benefits from being broken over multiple lines. There is still a chance that the loop increment could be blank, so consider that
                // and either break over two or three lines.
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(
                        "for ({0};",
                        loopVarInitialiser
                    ),
                    indentationDepth
                ));
                translationResult = translationResult.Add(new TranslatedStatement(
                    continuationCondition + ";" + ((loopIncrementWithLeadingSpaceIfNonBlank == "") ? ")" : ""),
                    indentationDepth + 1
                ));
                if (loopIncrementWithLeadingSpaceIfNonBlank != "")
                {
                    translationResult = translationResult.Add(new TranslatedStatement(
                        loopIncrementWithLeadingSpaceIfNonBlank.Trim() + ")",
                        indentationDepth + 1
                    ));
                }
            }
            translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
            var earlyExitNameIfAny = GetEarlyExitNameIfRequired(forBlock, scopeAccessInformation);
            if (earlyExitNameIfAny != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format("var {0} = false;", earlyExitNameIfAny.Name),
                    indentationDepth + 1
                ));
            }
            translationResult = translationResult.Add(
                Translate(forBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, earlyExitNameIfAny, indentationDepth + 1)
            );
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                // If error-trapping is enabled and loop constaint evaluation failed, then the loop is entered once but the loop variable value is not
                // altered and continuation criteria are never checked - instead, the loop exits after one pass in these circumstances
                translationResult = translationResult
                    .Add(new TranslatedStatement("if (!" + constraintsInitialisedFlagNameIfAny.Name + ")", indentationDepth + 1))
                    .Add(new TranslatedStatement("break;", indentationDepth + 2));
            }
            translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth));
            }
            if (guardClauseLines.Any())
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }
            foreach (var undeclaredVariable in undeclaredVariableReferencesAccessedByLoopConstraints)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
            var earlyExitFlagNamesToCheck = scopeAccessInformation.StructureExitPoints
                .Where(e => e.ExitEarlyBooleanNameIfAny != null)
                .Select(e => e.ExitEarlyBooleanNameIfAny.Name);
            if (earlyExitFlagNamesToCheck.Any())
            {
                // Perform early-exit checks for any scopeAccessInformation.StructureExitPoints - if this is FOR loop inside a DO..LOOP loop and an
                // EXIT DO was encountered within the FOR that must refer to the containing DO, then the FOR loop will have been broken out of, but
                // also a flag set that means that we must break further to get out of the DO loop.
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        "if (" + string.Join(" || ", earlyExitFlagNamesToCheck) + ")",
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement(
                        "break;",
                        indentationDepth + 1
                    ));
            }
            return translationResult.AddUndeclaredVariables(undeclaredVariableReferencesAccessedByLoopConstraints);
		}

        private NumericValueToken TryToGetExpressionAsNumericConstant(Expression expression)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");

            var tokens = expression.Tokens.ToArray();
            if (tokens.Length != 1)
                return null;
            return tokens[0] as NumericValueToken;
        }

        private CSharpName GetEarlyExitNameIfRequired(ForBlock forBlock, ScopeAccessInformation scopeAccessInformation)
        {
            if (forBlock == null)
                throw new ArgumentNullException("forBlock");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            if (!forBlock.ContainsLoopThatContainsMismatchedExitThatMustBeHandledAtThisLevel())
                return null;

            return _tempNameGenerator(new CSharpName("exitFor"), scopeAccessInformation);
        }

        private TranslationResult Translate(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, CSharpName earlyExitNameIfAny, int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // Add a StructureExitPoint entry for the current loop so that the "early-exit" logic described in the Translate method above is possible
            return base.TranslateCommon(
                base.GetWithinFunctionBlockTranslators(),
				blocks,
				scopeAccessInformation.AddStructureExitPoints(
                    earlyExitNameIfAny,
                    ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions.For
                ),
				indentationDepth
			);
		}
    }
}
