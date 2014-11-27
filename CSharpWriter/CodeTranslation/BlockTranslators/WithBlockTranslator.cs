using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using CSharpWriter.CodeTranslation.Extensions;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class WithBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public WithBlockTranslator(
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

		public TranslationResult Translate(WithBlock withBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (withBlock == null)
                throw new ArgumentNullException("withBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var translatedTargetReference = _statementTranslator.Translate(withBlock.Target, scopeAccessInformation, ExpressionReturnTypeOptions.Reference);
            var undeclaredVariables = translatedTargetReference.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(_nameRewriter.GetMemberAccessTokenName(v), _nameRewriter));
            foreach (var undeclaredVariable in undeclaredVariables)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

            var targetName = base._tempNameGenerator(new CSharpName("with"), scopeAccessInformation);
            var withBlockContentTranslationResult = Translate(
                withBlock.Content.ToNonNullImmutableList(),
                new ScopeAccessInformation(
                    scopeAccessInformation.Parent,
                    scopeAccessInformation.ScopeDefiningParent,
                    scopeAccessInformation.ParentReturnValueNameIfAny,
                    scopeAccessInformation.ErrorRegistrationTokenIfAny,
                    new ScopeAccessInformation.DirectedWithReferenceDetails(
                        targetName,
                        withBlock.Target.Tokens.First().LineIndex
                    ),
                    scopeAccessInformation.ExternalDependencies,
                    scopeAccessInformation.Classes,
                    scopeAccessInformation.Functions,
                    scopeAccessInformation.Properties,
                    scopeAccessInformation.Variables,
                scopeAccessInformation.StructureExitPoints
                ),
                indentationDepth
            );
            return new TranslationResult(
                withBlockContentTranslationResult.TranslatedStatements
                    .Insert(
                        new TranslatedStatement(
                            string.Format(
                                "var {0} = {1};",
                                targetName.Name,
                                translatedTargetReference.TranslatedContent
                            ),
                            indentationDepth
                        ),
                        0
                    ),
                withBlockContentTranslationResult.ExplicitVariableDeclarations,
                withBlockContentTranslationResult.UndeclaredVariablesAccessed.AddRange(undeclaredVariables)
            );
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
