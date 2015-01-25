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
            // values. Note that these temporary values are NOT re-evaluated each loop since this is how VBScript (unlike some other languages) work. Also
            // note the loopStart is generated after loopEnd and loopStep - this is because it may depend upon them. In "FOR i = 1 TO 10", loopStart will be
            // a VBScript "Integer" (Int16) but in "FOR i = 1 TO 32768" loopStart needs to be a VBScript "Long" (Int32) since VBScript tries to arrange for
            // the loop variable to be of a type that doesn't need to change each iteration. (It's possible for code within the loop to change the loop
            // variable's type, but that's another story).
            var undeclaredVariableReferencesAccessedByLoopConstraints = new NonNullImmutableList<NameToken>();
            var loopConstraintInitialisersWhereRequired = new List<LoopConstraintInitialiser>();
            string loopEnd;
            var numericLoopEndValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopTo);
            if (numericLoopEndValueIfAny != null)
            {
                // Note: We need to use the NumericValueToken's AsCSharpValue() method here when generating the translated "loopEnd" output since its
                // type might be important when determining what the type for "loopStart" will be - eg. the end value in "FOR i = 1 TO 20.2" results
                // in the loop variable being a double.
                loopEnd = numericLoopEndValueIfAny.AsCSharpValue();
            }
            else
            {
                // Same logic as for the loopStart value above applies here
                var loopEndExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopTo,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );
                var loopEndName = _tempNameGenerator(new CSharpName("loopEnd"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(new LoopConstraintInitialiser(
                    loopEndName,
                    string.Format(
                        "{0}.CDBL({1})",
                        _supportRefName.Name,
                        loopEndExpressionContent.TranslatedContent
                    ),
                    "double"
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
            {
                // Note: We need to use the NumericValueToken's AsCSharpValue() method here when generating the translated "loopStep" output since its
                // type might be important when determining what the type for "loopStart" will be - eg. the loop step in "FOR i = 1 TO 10 STEP 0.1"
                // results in the loop variable being a double.
                loopStep = numericLoopStepValueIfAny.AsCSharpValue();
            }
            else
            {
                // Same logic as for the loopStart/loopEnd value above applies here
                var loopStepExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopStep,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );
                var loopStepName = _tempNameGenerator(new CSharpName("loopStep"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(new LoopConstraintInitialiser(
                    loopStepName,
                    string.Format(
                        "{0}.CDBL({1})",
                        _supportRefName.Name,
                        loopStepExpressionContent.TranslatedContent
                    ),
                    "double"
                ));
                loopStep = loopStepName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    loopStepExpressionContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            // The loopStart value is what the loop variable will first be set to. This means that it has to follow some complicate VBScript logic. The
            // VBScript interpreter tries to ensure that a loop variable's type will not change during an iteration (unless something crazy within in
            // loop changes it, which is allowed to happen - but not something we need to worry about here). So in the loop "FOR i = 1 To 5" it's fine
            // for the loopStart to be a VBScript "Integer" (Int16) since that can run the loop from 1..5 without changing type. However, in the loop
            // "FOR i = 1 To 32768" an Int16 wouldn't be able to describe all of the values, 32768 would be an overflow. So VBScript would set the
            // loop variable to be a "Long" (Int32). It's important that the loopStart expression we generate here includes sufficient type information
            // to do the same sort of thing. (There is similar-but-different handling for dates - eg. "FOR i = Date() TO Date() + 2" - where the type
            // of the loop variable will be a Date, but if the loopEnd expression is too high - eg. "FOR i = Date() To 10000000" - then there will be
            // an overflow error, rather than the loop variable type being changed so that it can contain the entire range).
            string loopStart;
            var numericLoopStartValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopFrom);
            if ((numericLoopStartValueIfAny != null) && numericLoopStartValueIfAny.IsVBScriptInteger()
            && (numericLoopEndValueIfAny != null) && numericLoopEndValueIfAny.IsVBScriptInteger()
            && (numericLoopStepValueIfAny != null) && numericLoopStepValueIfAny.IsVBScriptInteger())
            {
                // For really simple loops, like "FOR i = 1 TO 5", where the loop constraints are all compile-time constants and all within the "Integer"
                // range, bypass the logic that performs the other checks and just call numericLoopStartValueIfAny.AsCSharpValue() - this will include
                // type information in the translated content (eg. "(Int16)1").
                // - One the one hand, it seems a bit silly bothering layering on more special cases when VBScript already has a million special cases
                //   of its own, but the exception here makes the translated output slightly less WTF-worthy for simple and common cases
                loopStart = numericLoopStartValueIfAny.AsCSharpValue();
            }
            else
            {
                // When determining what type the loop variable should be (eg. Int16 in "FOR i = 1 to 5" or Int32 in "FOR i = 1 TO 32768" or Double in
                // "FOR i = 1 TO 10 STEP 0.1"), the loop end and loop step values may affect the loop start value. However, if the loop end and step
                // values are compile-time numeric constants that are known to be of type VBScript "Integer" (Int16) then they won't affect the
                // loop variable type at all, so they needn't be considered.
                var numericValuesTheTypeMustBeAbleToContain = new List<string>();
                if ((numericLoopEndValueIfAny == null) || !numericLoopEndValueIfAny.IsVBScriptInteger())
                    numericValuesTheTypeMustBeAbleToContain.Add(loopEnd);
                if ((numericLoopStepValueIfAny == null) || !numericLoopStepValueIfAny.IsVBScriptInteger())
                    numericValuesTheTypeMustBeAbleToContain.Add(loopStep);

                var loopStartExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopFrom,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );

                var loopStartName = _tempNameGenerator(new CSharpName("loopStart"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(new LoopConstraintInitialiser(
                    loopStartName,
                    string.Format(
                        "{0}.NUM({1}{2}{3})",
                        _supportRefName.Name,
                        loopStartExpressionContent.TranslatedContent,
                        numericValuesTheTypeMustBeAbleToContain.Any() ? ", " : "",
                        string.Join(", ", numericValuesTheTypeMustBeAbleToContain)
                    ),
                    "object"
                ));
                loopStart = loopStartName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    loopStartExpressionContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            // Any dynamic loop constraints (ie. those that can be confirmed to be fixed numeric values at translation time) need to have variables
            // declared and initialised. There may be a mix of dynamic and constant constraints so there may be zero, one, two or three variables
            // to deal with here (this will have been determined in the work above).
            // - Note: The loopEnd and loopStep fixed constraints are stored as as doubles (where error-trapping is involved, the variables are explicitly
            //   declared as double, though otherwise var is used when CDBL is called). This means that the loop termination conditions are simpler, since
            //   we only need to call NUM on the loop variable - which may be set to anything by code within the loop - we can rest safe in the knowledge
            //   that these constraints are always numeric since they are doubles. However, loopStart has to be an object since it could be any of the
            //   VBScript "numeric" types and its type is what the loop variable is assigned when the loop begins; so if the "from" value is a DateTime
            //   then the loop variable starts as a DateTime and is incremented as a DateTime each iteration. So guard clauses involving loopStart
            //   need to use support functions such as StrictLTE.
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
                                loopConstraintInitialiser.VariableName.Name,
                                loopConstraintInitialiser.InitialisationContent
                            ),
                            indentationDepth
                        ));
                }
            }
            else
            {
                constraintsInitialisedFlagNameIfAny = _tempNameGenerator(new CSharpName("loopConstraintsInitialised"), scopeAccessInformation);
                foreach (var loopConstraint in loopConstraintInitialisersWhereRequired)
                {
                    translationResult = translationResult
                        .Add(new TranslatedStatement(
                            string.Format(
                                "{0} {1} = 0;",
                                loopConstraint.TypeNameIfPostponingSetting,
                                loopConstraint.VariableName.Name
                            ),
                            indentationDepth
                        ));
                }
                translationResult = translationResult
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
                                loopConstraintInitialiser.VariableName.Name,
                                loopConstraintInitialiser.InitialisationContent
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
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    // Ascending loop or infinite loop (step zero, which is supported in VBScript), start must not be greater than end
                    guardClauseLines = guardClauseLines.Add(string.Format("{0}.StrictLTE({1}, {2})", _supportRefName.Name, loopStart, loopEnd));
                }
                else
                {
                    // Descending loop, start must be greater than end
                    guardClauseLines = guardClauseLines.Add(string.Format("{0}.StrictGT({1} > {2})", _supportRefName.Name, loopStart, loopEnd));
                }
            }
            else if ((numericLoopStartValueIfAny != null) && (numericLoopEndValueIfAny != null))
            {
                // If the from and to bounds are known to be constants, but not the step then we just need to ensure that the step matches the
                // direction of from and to at runtime. Note: A step of zero will cause an infinite loop, but only if from <= to (the loop will
                // not be executed if it is descending and has a step of zero)
                if (numericLoopStartValueIfAny.Value <= numericLoopEndValueIfAny.Value)
                    guardClauseLines = guardClauseLines.Add(string.Format("({0} >= 0)", loopStep));
                else
                    guardClauseLines = guardClauseLines.Add(string.Format("({0} < 0)", loopStep));
            }
            else
            {
                // There are no more shortcuts now, we need to check at runtime that loopStep is negative for a descending loop and non-negative
                // for a non-descending loop
                guardClauseLines = guardClauseLines.Add(string.Format(
                    "(({0}.StrictLTE({1}, {2}) && ({3} >= 0))",
                    _supportRefName.Name,
                    loopStart,
                    loopEnd,
                    loopStep
                ));
                guardClauseLines = guardClauseLines.Add(string.Format(
                    "|| ({0}.StrictGT({1}, {2}) && ({3} < 0)))",
                    _supportRefName.Name,
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
                    // Note: If loopEnd is a known numeric constant then we can render just its value (eg. "1") instead of its translated content, which
                    // would include type information (eg. "(Int16)1") since this will make the final output a little more succinct (and the type information
                    // that we're skipping will have no effect on the StrictLTE call). This is most obvious for really simple loops but it doesn't hurt to
                    // try to make the more complicated ones shorter (since they get increasingly complicated and verbose as dynamic loop constraints and
                    // potential error-trapping are added!).
                    continuationCondition = string.Format(
                        "{0}.StrictLTE({1}, {2})",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        (numericLoopEndValueIfAny == null) ? loopEnd : numericLoopEndValueIfAny.Value.ToString()
                    );
                }
                else
                {
                    // Note: If loopEnd is a known numeric constant then we can render just its value instead of its translated content - see note above
                    continuationCondition = string.Format(
                        "{0}.StrictGTE({1}, {2})",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        (numericLoopEndValueIfAny == null) ? loopEnd : numericLoopEndValueIfAny.Value.ToString()
                    );
                }
            }
            else
            {
                // Note: If loopEnd is a known numeric constant then we can render just its value instead of its translated content - see note above
                continuationCondition = string.Format(
                    "(({3} >= 0) && {0}.StrictLTE({1}, {2})) || (({3} < 0) && {0}.StrictGTE({1}, {2}))",
                    _supportRefName.Name,
                    rewrittenLoopVariableName,
                    (numericLoopEndValueIfAny == null) ? loopEnd : numericLoopEndValueIfAny.Value.ToString(),
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
                        numericLoopStepValueIfAny.GetNegative().Value,
                        _supportRefName.Name
                    );
                }
                else
                {
                    // Another shortcut we can take is if the step value is known to be a non-negative numeric constant, we can just render its value
                    // out, rather than its full translated content (eg. "1" compared to "(Int16)1", since its type will not have any effect on the
                    // ADD action and rendering its numeric value only is more succinct)
                    loopIncrementWithLeadingSpaceIfNonBlank = string.Format(
                        " {0} = {2}.ADD({0}, {1})",
                        rewrittenLoopVariableName,
                        (numericLoopStepValueIfAny != null) ? numericLoopStepValueIfAny.Value.ToString() : loopStep,
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

        private class LoopConstraintInitialiser
        {
            public LoopConstraintInitialiser(CSharpName variableName, string initialisationContent, string typeNameIfPostponingSetting)
            {
                if (variableName == null)
                    throw new ArgumentNullException("variableName");
                if (string.IsNullOrWhiteSpace(initialisationContent))
                    throw new ArgumentException("Null/blank initialisationContent specified");
                if (string.IsNullOrWhiteSpace(typeNameIfPostponingSetting))
                    throw new ArgumentException("Null/blank typeNameIfPostponingSetting specified");

                VariableName = variableName;
                InitialisationContent = initialisationContent;
                TypeNameIfPostponingSetting = typeNameIfPostponingSetting;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public CSharpName VariableName { get; private set; }

            /// <summary>
            /// This will never be null or blank
            /// </summary>
            public string InitialisationContent { get; private set; }

            /// <summary>
            /// This will never be null or blank
            /// </summary>
            public string TypeNameIfPostponingSetting { get; private set; }
        }
    }
}
