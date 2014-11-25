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
            if (logger == null)
                throw new ArgumentNullException("logger");

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

            // Identify tokens for the start, end and step variables. If they are numeric constants then use them, otherwise they must be stored in temporary
            // values. Note that these temporary values are NOT re-evaluated each loop since this is how VBScript (unlike some other languages) work.
            var undeclaredVariableReferencesAccessedByLoopConstraints = new NonNullImmutableList<NameToken>();
            var loopConstraintInitialisersWhereRequired = new List<Tuple<CSharpName, string>>();
            string loopStart;
            var numericLoopStartValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopFrom);
            if (numericLoopStartValueIfAny != null)
                loopStart = numericLoopStartValueIfAny.Value.ToString();
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
                        "{0}.NUM({1})",
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
                loopEnd = numericLoopEndValueIfAny.Value.ToString();
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
                        "{0}.NUM({1})",
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
                ? new NumericValueToken(1, forBlock.LoopTo.Tokens.Last().LineIndex) // Default to Step 1 if no LoopStep expression was specified
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
                    numericLoopStepValueIfAny = new NumericValueToken(-lastLoopStepTokenAsNumericValueToken.Value, lastLoopStepTokenAsNumericValueToken.LineIndex);
            }
            if (numericLoopStepValueIfAny != null)
                loopStep = numericLoopStepValueIfAny.Value.ToString();
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
                        "{0}.NUM({1})",
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
                            "double {0};",
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
                            "{0}.HANDLEERROR({1}, () =>",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement("{", indentationDepth));
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
            string guardClause;
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
                guardClause = null;
            }
            else if (numericLoopStepValueIfAny != null)
            {
                // If the step is a known numeric constant, then the guard clause only needs to compare the from and to constaints (at least
                // one of which must not be a known constant, otherwise we'd be in the above condition)
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    // Ascending loop or infinite loop (step zero, which is supported in VBScript), start must not be greater than end
                    guardClause = string.Format("({0} <= {1})", loopStart, loopEnd);
                }
                else
                {
                    // Descending loop, start must be greater than end
                    guardClause = string.Format("({0} > {1})", loopStart, loopEnd);
                }
            }
            else if ((numericLoopStartValueIfAny != null) && (numericLoopEndValueIfAny != null))
            {
                // If the from and to bounds are known to be constants, but not the step then we just need to ensure that the step matches the
                // direction of from and to at runtime. Note: A step of zero will cause an infinite loop, but only if from <= to (the loop will
                // not be executed if it is descending and has a step of zero)
                if (numericLoopStartValueIfAny.Value <= numericLoopEndValueIfAny.Value)
                    guardClause = string.Format("({0} >= 0)", loopStep);
                else
                    guardClause = string.Format("({0} < 0)", loopStep);
            }
            else
            {
                // There are no more shortcuts now, we need to check at runtime that loopStep is negative for a descending loop and non-negative
                // for a non-descending loop
                guardClause = string.Format(
                    "((({0} <= {1}) && ({2} >= 0)) || (({0} > {1}) && ({2} < 0)))",
                    loopStart,
                    loopEnd,
                    loopStep
                );
            }

            // If there are non-constant constraint(s) and error-trapping may be enabled, then ensure there is a guard clause that ensures that
            // the constraints were successfully evalulated before considering entering the loop
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                guardClause = string.Format(
                    (guardClause == null) ? "({0})" : "({0} && {1})",
                    constraintsInitialisedFlagNameIfAny.Name,
                    guardClause
                );
            }

            var rewrittenLoopVariableName = _nameRewriter.GetMemberAccessTokenName(forBlock.LoopVar);
            var targetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenLoopVariableName, _envRefName, _outerRefName, _nameRewriter);
            if (targetContainer != null)
                rewrittenLoopVariableName = targetContainer.Name + "." + rewrittenLoopVariableName;

            if (guardClause != null)
            {
                translationResult = translationResult
                    .Add(new TranslatedStatement("if " + guardClause, indentationDepth))
                    .Add(new TranslatedStatement("{", indentationDepth));
                indentationDepth++;
            }
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "{0}.HANDLEERROR({1}, () =>",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement("{", indentationDepth));
                indentationDepth++;
            }
            string continuationCondition;
            if (numericLoopStepValueIfAny != null)
            {
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    continuationCondition = string.Format(
                        "{0}.NUM({1}) <= {2}",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        loopEnd
                    );
                }
                else
                {
                    continuationCondition = string.Format(
                        "{0}.NUM({1}) >= {2}",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        loopEnd
                    );
                }
            }
            else
            {
                continuationCondition = string.Format(
                    "(({3} >= 0) && ({0}.NUM({1}) <= {2})) || (({3} < 0) && ({0}.NUM({1}) >= {2}))",
                    _supportRefName.Name,
                    rewrittenLoopVariableName,
                    loopEnd,
                    loopStep
                );
            }
            string loopIncrementWithLeadingSpaceIfNonBlank;
            if ((numericLoopStepValueIfAny != null) && (numericLoopStepValueIfAny.Value == 0))
                loopIncrementWithLeadingSpaceIfNonBlank = "";
            else
            {
                loopIncrementWithLeadingSpaceIfNonBlank = string.Format(
                    " {0} = {3}.NUM({0}) {1} {2}",
                    rewrittenLoopVariableName,
                    ((numericLoopStepValueIfAny == null) || (numericLoopStepValueIfAny.Value >= 0)) ? "+" : "-",
                    ((numericLoopStepValueIfAny == null) || (numericLoopStepValueIfAny.Value >= 0)) ? loopStep : Math.Abs(numericLoopStepValueIfAny.Value).ToString(),
                    _supportRefName.Name
                );
            }
            translationResult = translationResult.Add(new TranslatedStatement(
                string.Format(
                    "for ({0} = {1}; {2};{3})",
                    rewrittenLoopVariableName,
                    loopStart,
                    continuationCondition,
                    loopIncrementWithLeadingSpaceIfNonBlank
                ),
                indentationDepth
            ));
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
            translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth));
            }
            if (guardClause != null)
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
