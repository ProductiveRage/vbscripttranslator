using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class IfBlockTranslator : CodeBlockTranslator
    {
        private readonly ILogInformation _logger;
        public IfBlockTranslator(
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

		public TranslationResult Translate(IfBlock ifBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (ifBlock == null)
                throw new ArgumentNullException("ifBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			var translationResult = TranslationResult.Empty;
            foreach (var conditionalEntry in ifBlock.ConditionalClauses.Select((conditional, index) => new { Conditional = conditional, Index = index }))
            {
                var conditionalContent = _statementTranslator.Translate(
                    conditionalEntry.Conditional.Condition,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );
                if (!scopeAccessInformation.MayRequireErrorWrapping(ifBlock))
                {
                    conditionalContent = new TranslatedStatementContentDetails(
                        string.Format(
                            "{0}.IF({1})",
                            _supportRefName.Name,
                            conditionalContent.TranslatedContent
                        ),
                        conditionalContent.VariablesAccessed
                    );
                }
                else
                {
                    conditionalContent = new TranslatedStatementContentDetails(
                        string.Format(
                            "{0}.IF(() => {1}, {2})",
                            _supportRefName.Name,
                            conditionalContent.TranslatedContent,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        conditionalContent.VariablesAccessed
                    );
                }
                var undeclaredVariablesAccessed = conditionalContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
                foreach (var undeclaredVariable in undeclaredVariablesAccessed)
                    _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesAccessed);
                translationResult = translationResult.Add(
                    new TranslatedStatement(
                        string.Format(
                            "{0} ({1})",
                            (conditionalEntry.Index == 0) ? "if" : "else if",
                            conditionalContent.TranslatedContent
                        ),
                        indentationDepth
                    )
                );
                translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                translationResult = translationResult.Add(
                    Translate(conditionalEntry.Conditional.Statements.ToNonNullImmutableList(), scopeAccessInformation, indentationDepth + 1)
                );
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }
            if (ifBlock.OptionalElseClause != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth));
                translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                translationResult = translationResult.Add(
                    Translate(ifBlock.OptionalElseClause.Statements.ToNonNullImmutableList(), scopeAccessInformation, indentationDepth + 1)
                );
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
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
