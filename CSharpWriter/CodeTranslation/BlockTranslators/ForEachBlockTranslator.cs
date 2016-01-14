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

			// The approach here is to get an IEnumerator reference and then loop over it in a "while (true)" loop, exiting when there are no more items. It would
			// feel more natural to use a C# foreach loop but the loop variable may not be restricted in scope to the loop (in fact, in VBScript this is very unlikely)
			// and so a "Type and identifier are both required in a foreach statement" compile error would result - "foreach (i in a)" is not valid, it must be of the
			// form "foreach (var i in a)" which only works if "i" is limited in scope to that loop. The "while (true)" structure also works better when error-trapping
			// may be enabled since the call-MoveNext-and-set-loop-variable-to-enumerator-Current-value-if-not-reached-end-of-data can be bypassed entirely if an error
			// was caught while evaluating the enumerator (in which case the loop should be processed once but the loop variable not set).

			// Note: The looped-over content must be of type "Reference" since VBScript won't enumerate over strings, for example, whereas C# would be happy to.
			// However, to make the output marginally easier to read, the ENUMERABLE method will deal with this logic and so the ExpressionReturnTypeOptions
			// value passed to the statement translator is "NotSpecified".

			var loopSourceContent = _statementTranslator.Translate(forEachBlock.LoopSrc, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
			var undeclaredVariablesInLoopSourceContent = loopSourceContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
			foreach (var undeclaredVariable in undeclaredVariablesInLoopSourceContent)
				_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

			var translationResult = TranslationResult.Empty.AddUndeclaredVariables(undeclaredVariablesInLoopSourceContent);
			var enumerationContentVariableName = _tempNameGenerator(new CSharpName("enumerationContent"), scopeAccessInformation);
			var enumeratorInitialisationContent = string.Format(
				"{0} = {1}.ENUMERABLE({2}).GetEnumerator();",
				enumerationContentVariableName.Name,
				_supportRefName.Name,
				loopSourceContent.TranslatedContent
			);
			var rewrittenLoopVarName = _nameRewriter.GetMemberAccessTokenName(forEachBlock.LoopVar);
			var loopVarTargetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(forEachBlock.LoopVar, _envRefName, _outerRefName, _nameRewriter);
			if (loopVarTargetContainer != null)
				rewrittenLoopVarName = loopVarTargetContainer.Name + "." + rewrittenLoopVarName;
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
			{
				translationResult = translationResult.Add(new TranslatedStatement(
					"var " + enumeratorInitialisationContent,
					indentationDepth,
					forEachBlock.LoopVar.LineIndex
				));
			}
			else
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
				translationResult = translationResult
					.Add(new TranslatedStatement(
						string.Format(
							"IEnumerator {0} = null;",
							enumerationContentVariableName.Name
						),
						indentationDepth,
						forEachBlock.LoopVar.LineIndex
					))
					.Add(new TranslatedStatement(
						string.Format(
							"{0}.HANDLEERROR({1}, () => {{",
							_supportRefName.Name,
							scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
						),
						indentationDepth,
						forEachBlock.LoopVar.LineIndex
					))
					.Add(new TranslatedStatement(
						enumeratorInitialisationContent,
						indentationDepth + 1,
						forEachBlock.LoopVar.LineIndex
					))
					.Add(new TranslatedStatement("});", indentationDepth, forEachBlock.LoopVar.LineIndex));
			}
			translationResult = translationResult
				.Add(new TranslatedStatement("while (true)", indentationDepth, forEachBlock.LoopVar.LineIndex))
				.Add(new TranslatedStatement("{", indentationDepth, forEachBlock.LoopVar.LineIndex));
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				// If error-trapping is enabled and an error was indeed trapped while trying evaluate the enumerator, then the enumerator will be null.
				// In this case, the loop should be executed once but the loop variable not set to anything. When this happens, there is no point trying
				// to call MoveNext (since the enumerator is null) and the loop-variable-setting should be skipped. So an is-null check is wrapper around
				// that work. If error-trapping is not enabled then this check is not required and a level of nesting in the translated output can be
				// avoided.
				translationResult = translationResult
					.Add(new TranslatedStatement(
						string.Format(
							"if ({0} != null)",
							enumerationContentVariableName.Name
						),
						indentationDepth + 1,
						forEachBlock.LoopVar.LineIndex
					))
					.Add(new TranslatedStatement("{", indentationDepth + 1, forEachBlock.LoopVar.LineIndex));
				indentationDepth++;
			}
			translationResult = translationResult
				.Add(new TranslatedStatement(string.Format(
						"if (!{0}.MoveNext())",
						enumerationContentVariableName.Name
					),
					indentationDepth + 1,
					forEachBlock.LoopVar.LineIndex
				))
				.Add(new TranslatedStatement("break;", indentationDepth + 2, forEachBlock.LoopVar.LineIndex))
				.Add(new TranslatedStatement(
					string.Format(
						"{0} = {1}.Current;",
						rewrittenLoopVarName,
						enumerationContentVariableName.Name
					),
					indentationDepth + 1,
					forEachBlock.LoopVar.LineIndex
				));
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				// If error-trapping may be enabled then the above MoveNext and set-to-Current work was wrapped in a condition which must be closed
				translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth, forEachBlock.LoopVar.LineIndex));
				indentationDepth--;
			}
			var earlyExitNameIfAny = GetEarlyExitNameIfRequired(forEachBlock, scopeAccessInformation);
			if (earlyExitNameIfAny != null)
			{
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format("var {0} = false;", earlyExitNameIfAny.Name),
					indentationDepth + 1,
					forEachBlock.LoopVar.LineIndex
				));
			}
			translationResult = translationResult.Add(
				Translate(
					forEachBlock.Statements.ToNonNullImmutableList(),
					scopeAccessInformation.SetParent(forEachBlock),
					earlyExitNameIfAny,
					indentationDepth + 1
				)
			);
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				// If error-trapping was enabled and an error caught, then the loop should be processed once and only once. The enumerator reference
				// will be null - so check for that and exit if so. If there is no chance that error-trapping is enabled then this condition is not
				// required and there is no point emitting it.
				translationResult = translationResult
					.Add(new TranslatedStatement(
						string.Format(
							"if ({0} == null)",
							enumerationContentVariableName.Name
						),
						indentationDepth + 1,
						forEachBlock.LoopVar.LineIndex
					))
					.Add(new TranslatedStatement("break;", indentationDepth + 2, forEachBlock.LoopVar.LineIndex));
			}
			translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth, forEachBlock.LoopVar.LineIndex));

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
						indentationDepth,
						forEachBlock.LoopVar.LineIndex
					))
					.Add(new TranslatedStatement(
						"break;",
						indentationDepth + 1,
						forEachBlock.LoopVar.LineIndex
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
