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

            // Note: The looped-over content must be of type "Reference" since VBScript won't enumerate over strings, for example, whereas C# would be happy to.
            // However, to make the output marginally easier to read, the ENUMERABLE method will deal with this logic and so the ExpressionReturnTypeOptions
            // value passed to the statement translator is "NotSpecified".
            var loopSourceContent = _statementTranslator.Translate(forEachBlock.LoopSrc, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
            var undeclaredVariablesInLoopSourceContent = loopSourceContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
            foreach (var undeclaredVariable in undeclaredVariablesInLoopSourceContent)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

            var translationResult = TranslationResult.Empty;
            var enumerationContent = string.Format(
                "{0}.ENUMERABLE({1})",
                _supportRefName.Name,
                loopSourceContent.TranslatedContent
            );
            var rewrittenLoopVarName = _nameRewriter.GetMemberAccessTokenName(forEachBlock.LoopVar);
            var loopVarTargetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenLoopVarName, _envRefName, _outerRefName, _nameRewriter);
            if (loopVarTargetContainer != null)
                rewrittenLoopVarName = loopVarTargetContainer.Name + "." + rewrittenLoopVarName;
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                // If ON ERROR RESUME NEXT wraps a FOR EACH loop and there is an error in evaluating the enumerator, then the loop will be entered once. The
                // loop variable will not be altered - eg.
                //
                //   On Error Resume Next
                //   Dim i: For Each i in "12"
                //     WScript.Echo "We're in the loop! i is " & TypeName(i)
                //   Next
                //
                // VBScript can not enumerate a string, so the loop errors. But the ON ERROR RESUME NEXT causes the loop to be processed - only once. This is
                // approached by calling _.ENUMERABLE before the loop construct, setting a temporary variable to be the returned enumerable inside a call to
                // "HANDLEERROR" - meaning that it will be left as null if it fails. Null values are replaced with a single-item array, where the element is
                // the current value of the loop variable - so its value is not altered when the loop is entered.
                // - Note: The "error-trapping" wrapper functions ("HANDLEERROR") are used around code that MAY have error-trapping enabled (there are cases
                //   where we can't know at translation time whether it will have been turned off with ON ERROR GOTO 0 or not - and there are probably some
                //   cases that could be picked up if the translation process was more intelligent). If the above example had an ON ERROR GOTO 0 between the
                //   ON ERROR RESUME NEXT and the FOR EACH loop then the error (about trying to enumerate a string) would be raised and so the FOR EACH would
                //   not be entered, and so the single-element "fallback array" would never come to exist. If errors WERE still being captured, then the
                //   translated FOR EACH loop WOULD be entered and enumerated through once (without the value of the loop variable being altered, as
                //   is consistent with VBScript)
                var enumerationContentVariableName = _tempNameGenerator(new CSharpName("enumerationContent"), scopeAccessInformation);
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "IEnumerable {0} = null;",
                            enumerationContentVariableName.Name
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
                    ))
                    .Add(new TranslatedStatement(
                        string.Format(
                            "{0} = {1};",
                            enumerationContentVariableName.Name,
                            enumerationContent
                        ),
                        indentationDepth + 1
                    ))
                    .Add(new TranslatedStatement("});", indentationDepth));
                enumerationContent = enumerationContentVariableName.Name + " ?? new object[] { " + rewrittenLoopVarName + " }";
            }
            translationResult = translationResult
                .Add(new TranslatedStatement(
                    string.Format(
                        "foreach ({0} in {1})",
                        rewrittenLoopVarName,
                        enumerationContent
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
