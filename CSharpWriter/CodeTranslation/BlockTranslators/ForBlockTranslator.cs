using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
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

            var translationResult = TranslationResult.Empty;

            // Identify tokens for the start, end and step variables. If they are numeric constants then use them, otherwise they must be stored in temporary
            // values. Note that these temporary values are NOT re-evaluated each loop since this is how VBScript (unlike some other languages) work.
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
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "var {0} = {1}.NUM({2});",
                            loopStartName.Name,
                            _supportRefName.Name,
                            loopStartExpressionContent
                        ),
                        indentationDepth
                    ));
                loopStart = loopStartName.Name;
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
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "var {0} = {1}.NUM({2});",
                            loopEndName.Name,
                            _supportRefName.Name,
                            loopEndExpressionContent
                        ),
                        indentationDepth
                    ));
                loopEnd = loopEndName.Name;
            }
            string loopStep;
            var numericLoopStepValueIfAny = (forBlock.LoopStep == null)
                ? new NumericValueToken(1, forBlock.LoopTo.Tokens.Last().LineIndex) // Default to Step 1 if no LoopStep expression was specified
                : TryToGetExpressionAsNumericConstant(forBlock.LoopStep);
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
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "var {0} = {1}.NUM({2});",
                            loopStepName.Name,
                            _supportRefName.Name,
                            loopStepExpressionContent
                        ),
                        indentationDepth
                    ));
                loopStep = loopStepName.Name;
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

            // TODO: Comparison should be ">=" or "<=" (depending upon whether loopStart < loopEnd or not)
            // - Have simple case if loop start and end are fixed constants but require more verbose "OR" combination if one of both are not

            var rewrittenLoopVariableName = _nameRewriter(forBlock.LoopVar).Name;
            var targetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenLoopVariableName, _envRefName, _outerRefName, _nameRewriter);
            if (targetContainer != null)
                rewrittenLoopVariableName = targetContainer.Name + "." + rewrittenLoopVariableName;

            var indentationDepthLoop = indentationDepth;
            if (guardClause != null)
            {
                translationResult = translationResult
                    .Add(new TranslatedStatement("if " + guardClause, indentationDepth))
                    .Add(new TranslatedStatement("{", indentationDepth));
                indentationDepthLoop++;
            }
            translationResult = translationResult.Add(new TranslatedStatement(
                string.Format(
                    "for ({0} = {1}; {5}.NUM({0}) < {2}; {0} = {5}.NUM({0}) {3} {4})",
                    rewrittenLoopVariableName,
                    loopStart,
                    loopEnd,
                    ((numericLoopStepValueIfAny == null) || (numericLoopStepValueIfAny.Value >= 0)) ? "+" : "-",
                    ((numericLoopStepValueIfAny == null) || (numericLoopStepValueIfAny.Value >= 0)) ? loopStep : Math.Abs(numericLoopStepValueIfAny.Value).ToString(),
                    _supportRefName.Name
                ),
                indentationDepthLoop
            ));
            translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepthLoop));
            translationResult = translationResult.Add(
                Translate(forBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, indentationDepthLoop + 1)
            );
            translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepthLoop));
            if (guardClause != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }
            return translationResult;
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

		private TranslationResult Translate(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			return base.TranslateCommon(
				new BlockTranslationAttempter[]
				{
					base.TryToTranslateBlankLine,
					base.TryToTranslateComment,
					base.TryToTranslateDim,
					base.TryToTranslateDo,
					base.TryToTranslateExit,
					base.TryToTranslateFor,
					base.TryToTranslateForEach,
					base.TryToTranslateIf,
                    base.TryToTranslateOnErrorResumeNext,
                    base.TryToTranslateOnErrorGotoZero,
					base.TryToTranslateReDim,
					base.TryToTranslateRandomize,
					base.TryToTranslateStatementOrExpression,
					base.TryToTranslateSelect,
                    base.TryToTranslateValueSettingStatement
				}.ToNonNullImmutableList(),
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
		}

    }
}
