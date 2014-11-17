using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class DoBlockTranslator : CodeBlockTranslator
    {
        private readonly ILogInformation _logger;
        public DoBlockTranslator(
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

		public TranslationResult Translate(DoBlock doBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (doBlock == null)
                throw new ArgumentNullException("doBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            TranslatedStatementContentDetails whileConditionExpressionContentIfAny;
            if (doBlock.ConditionIfAny == null)
                whileConditionExpressionContentIfAny = null;
            else
            {
                whileConditionExpressionContentIfAny = _statementTranslator.Translate(
                    doBlock.ConditionIfAny,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.Boolean
                );
                if (!doBlock.IsDoWhileCondition)
                {
                    // C# doesn't support "DO UNTIL x" but it's equivalent to "DO WHILE !x"
                    whileConditionExpressionContentIfAny = new TranslatedStatementContentDetails(
                        "!" + whileConditionExpressionContentIfAny.TranslatedContent,
                        whileConditionExpressionContentIfAny.VariablesAccessed
                    );
                }
            }

            var translationResult = TranslationResult.Empty;
            if (whileConditionExpressionContentIfAny != null)
            {
                translationResult = translationResult.AddUndeclaredVariables(
                    whileConditionExpressionContentIfAny.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            if (whileConditionExpressionContentIfAny == null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "while (true)",
                    indentationDepth
                ));
            }
            else if (doBlock.IsPreCondition)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "do while (" + whileConditionExpressionContentIfAny.TranslatedContent + ")",
                    indentationDepth
                ));
            }
            else
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "do",
                    indentationDepth
                ));
            }
            translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
            translationResult = translationResult.Add(
                Translate(doBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, indentationDepth + 1)
            );
            if ((whileConditionExpressionContentIfAny == null) || doBlock.IsPreCondition)
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            else
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "} while (" + whileConditionExpressionContentIfAny.TranslatedContent + ");",
                    indentationDepth
                ));
            }
            return translationResult;
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
                base.GetWithinFunctionBlockTranslators(),
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
		}
    }
}
