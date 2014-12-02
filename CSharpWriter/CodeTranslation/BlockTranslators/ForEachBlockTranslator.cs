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
    public class ForEachBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public ForEachBlockTranslator(
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

		public TranslationResult Translate(ForEachBlock forEachBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (forEachBlock == null)
                throw new ArgumentNullException("forEachBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // The looped-over content must be of type "Reference" since VBScript won't enumerate over strings, for example, whereas C# would be happy to.
            // However, to make the output marginally easier to read, the ENUMERABLE method will deal with this logic and so the ExpressionReturnTypeOptions
            // value passed to the statement translator is "NotSpecified".
            var translationResult = TranslationResult.Empty;
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
            var loopSourceContent = _statementTranslator.Translate(forEachBlock.LoopSrc, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified);
            var undeclaredVariablesInLoopSourceContent = loopSourceContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
            foreach (var undeclaredVariable in undeclaredVariablesInLoopSourceContent)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
            var rewrittenLoopVarName = _nameRewriter.GetMemberAccessTokenName(forEachBlock.LoopVar);
            var loopVarTargetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenLoopVarName, _envRefName, _outerRefName, _nameRewriter);
            if (loopVarTargetContainer != null)
                rewrittenLoopVarName = loopVarTargetContainer.Name + "." + rewrittenLoopVarName;
            translationResult = translationResult
                .Add(new TranslatedStatement(
                    string.Format(
                        "for each ({0} in {1}.ENUMERABLE({2})",
                        rewrittenLoopVarName,
                        _supportRefName.Name,
                        loopSourceContent.TranslatedContent
                    ),
                    indentationDepth
                ))
                .AddUndeclaredVariables(undeclaredVariablesInLoopSourceContent);
            translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
            var earlyExitNameIfAny = GetEarlyExitNameIfRequired(forEachBlock, scopeAccessInformation);
            if (earlyExitNameIfAny != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format("var {0} = false;", earlyExitNameIfAny.Name),
                    indentationDepth + 1
                ));
            }
            translationResult = translationResult.Add(
                Translate(forEachBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, earlyExitNameIfAny, indentationDepth + 1)
            );
            translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth));
            }
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
            return translationResult;
		}

        private CSharpName GetEarlyExitNameIfRequired(ForEachBlock forEachBlock, ScopeAccessInformation scopeAccessInformation)
        {
            if (forEachBlock == null)
                throw new ArgumentNullException("forEachBlock");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            if (!forEachBlock.ContainsLoopThatContainsMismatchedExitThatMustBeHandledAtThisLevel())
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
