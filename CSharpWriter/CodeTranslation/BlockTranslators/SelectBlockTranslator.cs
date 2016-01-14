using System;
using System.Collections.Generic;
using System.Linq;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
	// TODO: Test..
	//  1. Literal value cases enforce appropriate comparisons (strings, dates, numbers?)
	//  2. Literal target enforces appropriate comparisons, where required
	//  3. VAL is not called on the target (needs to be called for each comparison)

	public class SelectBlockTranslator : CodeBlockTranslator
	{
		private readonly ITranslateIndividualStatements _statementTranslator;
		private readonly ILogInformation _logger;
		public SelectBlockTranslator(
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

		public TranslationResult Translate(SelectBlock selectBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (selectBlock == null)
				throw new ArgumentNullException("selectBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			// Notes:
			// 1. Case values are lazily evaluated; as soon as one is matched, no more are considered.
			// 2. There is no fall-through from one section to another; the first matched section (if any) is processed and no more are considered.

			var translationResult = TranslationResult.Empty;
			foreach (var openingComment in selectBlock.OpeningComments)
				translationResult = base.TryToTranslateComment(translationResult, openingComment, scopeAccessInformation, indentationDepth);

			// Do all of the work to decide what needs doing with the target expression (if it's not a simple value then it will be evaluating - this is done once for the entire
			// block and the resulting value reused for each case-value comparison)
			var targetExpressionTranslationDetails = TranslateTargetExpression(selectBlock.Expression, translationResult, scopeAccessInformation, indentationDepth);
			translationResult = targetExpressionTranslationDetails.ExtendedTranslationResult;
			scopeAccessInformation = targetExpressionTranslationDetails.ExtendedScopeAccessInformation;

			// If evaluation of the target expression fails at runtime then no comparisons work is required - note that this check is not required if there are zero comparisons
			// to deal with. In this case, VBScript would still evaluate the target even though there's nothing to compare it to (and we remain consistent with that; the target
			// evaluation occurs above but we do nothing more if there are no comparison values).
			if ((targetExpressionTranslationDetails.SuccessfullyEvaluatedTargetNameIfRequired != null) && selectBlock.Content.Any())
			{
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format("if ({0})", targetExpressionTranslationDetails.SuccessfullyEvaluatedTargetNameIfRequired.Name),
					indentationDepth,
					selectBlock.Expression.Tokens.First().LineIndex
				));
				translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth, selectBlock.Expression.Tokens.First().LineIndex));
				indentationDepth++;
			}

			// In simple cases, a SELECT CASE structure can be represented by a set of "if { .. } else if { .. } else if { .. } else { .. }" configuration. However, if we have to
			// do any by-ref argument aliasing (if we're in a function with by-ref arguments and any of those arguments is required within case-value-matching AND that case-value-
			// matching uses lambdas to wrap the work; if error-trapping is enabled, generally) then we need to use temporary booleans for each "CASE" did-match result, which means
			// that the structure has to become progressively nested - eg. "if { .. } else { if { .. } else { if { .. } else { .. } } }". Each time a CASE is processed, it records
			// whether an additional level of nesting was required for it, in which case the next CASE will start with an "if" rather than an "else if". Obviously this should be
			// the case for the very first CASE that will be encountered. On a related note, the number of times that this is required is recorded since when everything else is
			// complete, there will need to be additional closing braces.
			var shouldNextCaseValueBlockBeFreshIfBlock = true;
			var numberOfAdditionalIndentsRequiredForByRefAliasing = 0;

			var annotatedCaseBlocks = selectBlock.Content.Select((c, i) =>
			{
				var lastIndex = selectBlock.Content.Count() - 1;
				return new
				{
					IsFirstBlock = (i == 0),
					IsCaseLastBlockWithValuesToCheck =
						((i == lastIndex) && (c is SelectBlock.CaseBlockExpressionSegment)) ||
						((i == (lastIndex - 1)) && (c is SelectBlock.CaseBlockExpressionSegment) && (selectBlock.Content.Last() is SelectBlock.CaseBlockElseSegment)),
					CaseBlock = c
				};
			});
			foreach (var annotatedCaseBlock in annotatedCaseBlocks)
			{
				var explicitOptionsCaseBlock = annotatedCaseBlock.CaseBlock as SelectBlock.CaseBlockExpressionSegment;
				if (explicitOptionsCaseBlock != null)
				{
					// For each case value (VBScript supports multiple values per case, unlike languages like JavaScript and C#). For each value, we'll be generating part of a
					// combined "if" condition. The first thing we need to do with these segments is log any undeclared variables that are accessed. We'll use this data later
					// on, too, to generate the "if" condition (since we'll often use this at least twice - see note about by-ref argument handling further down as to why we
					// don't ALWAYS use it twice - and since we will always access EVERY value while checking for undeclared variable access, we might as well ToArray it to
					// prevent repeating any work).
					var conditions = explicitOptionsCaseBlock.Values
						.Select((value, index) => GetComparison(
							targetExpressionTranslationDetails.EvaluatedTarget,
							value,
							isFirstValueInCaseSet: (index == 0),
							scopeAccessInformation: scopeAccessInformation
						))
						.ToArray();

					var undeclaredVariablesInCondition = conditions.SelectMany(c => c.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter));
					foreach (var undeclaredVariable in undeclaredVariablesInCondition)
						_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
					translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesInCondition);

					// We need to record line index values for the "scaffolding" C# code that will be emitted here - we'll approximate by taking the line index of the first
					// value in the case options. This might not be perfect, depending upon the formatting of the VBScript source, but it should be close enough,
					var lineIndexForScaffolding = explicitOptionsCaseBlock.Values.First().Tokens.First().LineIndex;

					// If we're inside a function with by-ref arguments and any of those arguments needs to be accessed within a lambda (such as if error-trapping is enabled,
					// meaning we'll be performing the "IF" comparisons using the method signature that takes the expression to evaluate in as a lambda) then we'll need aliases
					// for those by-ref arguments. In this case, we'll use a temporary boolean to record the did-any-case-values-match-the-target-expression and set this flag
					// using the aliases. We'll then tidy up the aliases and the "if" block itself will just reference that flag. If there are no by-ref argument aliases
					// required then this complexity and temporary boolean are not required.
					var byRefArgumentIdentifier = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
					var byRefArgumentsToRewrite = new NonNullImmutableList<FuncByRefMapping>();
					foreach (var caseValue in explicitOptionsCaseBlock.Values)
					{
						byRefArgumentsToRewrite = byRefArgumentIdentifier.GetByRefArgumentsThatNeedRewriting(
							caseValue.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
							scopeAccessInformation,
							byRefArgumentsToRewrite
						);
					}
					ConditionMatchingByRefArgAliasingDetails byRefArgAliasMappingDetailsIfRequired;
					if (byRefArgumentsToRewrite.Any())
					{
						// If this is not the first CASE block then we need to move in a level of indentation to limit the scope of the temporary boolean value but also to ensure
						// that if any earlier CASE values matched that these are not evaluated and considered (nor are any later CASE values)
						if (!annotatedCaseBlock.IsFirstBlock)
						{
							translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth, lineIndexForScaffolding));
							translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth, lineIndexForScaffolding));
							indentationDepth++;
							numberOfAdditionalIndentsRequiredForByRefAliasing++;
						}

						var isCaseMatchResultName = _tempNameGenerator(new CSharpName("isCaseMatch"), scopeAccessInformation);
						translationResult = translationResult.Add(new TranslatedStatement(
							"bool " + isCaseMatchResultName.Name + ";",
							indentationDepth,
							lineIndexForScaffolding
						));
						var byRefAliasWrappingDetailsIfAny = byRefArgumentsToRewrite.OpenByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
						translationResult = byRefAliasWrappingDetailsIfAny.TranslationResult;
						indentationDepth += byRefAliasWrappingDetailsIfAny.DistanceToIndentCodeWithMappedValues;
						scopeAccessInformation = scopeAccessInformation.ExtendVariables(
							byRefArgumentsToRewrite
								.Select(r => new ScopedNameToken(r.To.Name, r.From.LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith))
								.ToNonNullImmutableList()
						);
						byRefArgAliasMappingDetailsIfRequired = new ConditionMatchingByRefArgAliasingDetails(
							byRefArgumentsToRewrite,
							byRefAliasWrappingDetailsIfAny,
							isCaseMatchResultName
						);

						// The case-value-match condition fragments need to be rewritten to reference the aliases, rather than the by-ref arguments (this means that these fragments
						// are safe to use within lambdas now)
						conditions = explicitOptionsCaseBlock.Values
							.Select((value, index) => GetComparison(
								targetExpressionTranslationDetails.EvaluatedTarget,
								byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(value, _nameRewriter),
								isFirstValueInCaseSet: (index == 0),
								scopeAccessInformation: scopeAccessInformation
							))
							.ToArray();
					}
					else
					{
						// If the previous CASE block needed some by-ref alias jiggery pokery and required that this block be indented inside its own "else", then create that structure now
						if (!annotatedCaseBlock.IsFirstBlock && shouldNextCaseValueBlockBeFreshIfBlock)
						{
							translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth, lineIndexForScaffolding));
							translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth, lineIndexForScaffolding));
							indentationDepth++;
							numberOfAdditionalIndentsRequiredForByRefAliasing++;
						}
						byRefArgAliasMappingDetailsIfRequired = null;
					}

					// If byRefArgAliasMappingDetailsIfRequired is null, then there was no funny business - we can just generate an "if" block directly, there mustn't have been any by-ref
					// argument handling to do. On the other hand, if byRefArgAliasMappingDetailsIfRequired is NOT null then we need to evaluate the do-any-conditions-match-the-target
					// possibilities and store then in a temporary boolean - then the by-ref argument aliases need tidying up and THEN we open an "if" block, using the temporary flag.
					if (byRefArgAliasMappingDetailsIfRequired == null)
					{
						translationResult = OpenIfBlockDirectly(
							translationResult,
							indentationDepth,
							conditions,
							openAsElseIf: !shouldNextCaseValueBlockBeFreshIfBlock,
							errorRegistrationTokenIfAny: scopeAccessInformation.ErrorRegistrationTokenIfAny,
							lineIndex: explicitOptionsCaseBlock.Values.First().Tokens.First().LineIndex
						);
					}
					else
					{
						translationResult = SetCaseMatchResultValue(
							translationResult,
							indentationDepth,
							conditions,
							byRefArgAliasMappingDetailsIfRequired.CaseValueMatchResultName,
							scopeAccessInformation.ErrorRegistrationTokenIfAny,
							lineIndex: explicitOptionsCaseBlock.Values.First().Tokens.First().LineIndex
						);
						indentationDepth -= byRefArgAliasMappingDetailsIfRequired.ByRefAliasWrappingDetails.DistanceToIndentCodeWithMappedValues;
						translationResult = byRefArgumentsToRewrite.CloseByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
						translationResult = translationResult.Add(new TranslatedStatement(
							string.Format(
								"if ({0})",
								byRefArgAliasMappingDetailsIfRequired.CaseValueMatchResultName.Name
							),
							indentationDepth,
							lineIndexForScaffolding
						));
					}

					translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth, lineIndexForScaffolding));
					translationResult = translationResult.Add(
						Translate(
							annotatedCaseBlock.CaseBlock.Statements.ToNonNullImmutableList(),
							scopeAccessInformation.SetParent(selectBlock),
							indentationDepth + 1
						)
					);
					translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth, translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource));

					// If we needed to break up the simple "if { .. } else if { .. } else if { .. } else { .. }" structure because by-ref argument aliasing was required and a temporary
					// boolean created for this CASE block, then the next CASE block with values (if there is one) needs to start a new "if" block (and not try to pick up with an "else if")
					shouldNextCaseValueBlockBeFreshIfBlock = byRefArgumentsToRewrite.Any();
				}
				else
				{
					var defaultCaseIsTheOnlyCase = (selectBlock.Content.Count() == 1);
					if (!defaultCaseIsTheOnlyCase)
					{
						// We need to record line index values for the "scaffolding" C# code that will be emitted here - when dealing with a case that matches value(s) then we use the line
						// index of one of those values. Here, though, there are no values (it's the Default-Match). However, since we know here that that this is not the ONLY case option
						// then there must have been some translated content emitted above (since the SelectBlock does not allow any CASE block other than the last one to be Default), so
						// we can just borrow the previous line's line index.
						translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth, translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource));
						translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth, translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource));
						indentationDepth++;
					}
					translationResult = translationResult.Add(
						Translate(
							annotatedCaseBlock.CaseBlock.Statements.ToNonNullImmutableList(),
							scopeAccessInformation.SetParent(selectBlock),
							indentationDepth
						)
					);
					if (!defaultCaseIsTheOnlyCase)
					{
						indentationDepth--;
						translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth, translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource));
					}
				}
			}
			for (var i = 0; i < numberOfAdditionalIndentsRequiredForByRefAliasing; i++)
			{
				indentationDepth--;
				translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth, translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource));
			}

			// If error-trapping may be active at runtime then the meat of translated content will have been wrapped in an "if", based upon whether the select target was successfully evaluated (in which case
			// we'll need to close that content here). Note that this will not have been the case if there were no expression to compare the target to (VBScript allows this - it evaluates the target expression
			// but then does nothing more).
			if ((targetExpressionTranslationDetails.SuccessfullyEvaluatedTargetNameIfRequired != null) && selectBlock.Content.Any())
			{
				indentationDepth--;
				translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth, translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource));
			}

			return translationResult;
		}

		private TranslationResult OpenIfBlockDirectly(
			TranslationResult translationResult,
			int indentationDepth,
			IEnumerable<TranslatedStatementContentDetails> conditionSegments,
			bool openAsElseIf,
			CSharpName errorRegistrationTokenIfAny,
			int lineIndex)
		{
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth");
			if (conditionSegments == null)
				throw new ArgumentNullException("conditionSegments");
			if (lineIndex < 0)
				throw new ArgumentOutOfRangeException("Must be zero or greater", "lineIndex");

			var conditionSegmentsArray = conditionSegments.ToArray();
			if (conditionSegmentsArray.Length == 0)
				throw new ArgumentException("There must be at least one condition segment");
			if (conditionSegmentsArray.Any(segment => segment == null))
				throw new ArgumentException("Null reference encountered in conditionSegments set");

			var wrappedConditionSegments = conditionSegments
				.Select(segment => string.Format(
					(errorRegistrationTokenIfAny == null) ? "{0}.IF({1})" : "{0}.IF(() => {1}, {2})",
					_supportRefName.Name,
					segment.TranslatedContent,
					(errorRegistrationTokenIfAny == null) ? "" : errorRegistrationTokenIfAny.Name
				))
				.ToArray();

			// The conditions where if blocks are opened directly (no by-ref argument aliasing and no error-trapping) tend to be quite short in a lot of cases - if the combined
			// content is not that long then put them all on a single line
			if (wrappedConditionSegments.Sum(segment => segment.Length) < 80)
			{
				return translationResult.Add(new TranslatedStatement(
					string.Format(
						"{0}if ({1})",
						openAsElseIf ? "else " : "",
						string.Join(" || ", wrappedConditionSegments.Select(segment => segment))
					),
					indentationDepth,
					lineIndex
				));
			}

			// Otherwise generate one line per condition
			for (var i = 0; i < wrappedConditionSegments.Length; i++)
			{
				string format;
				if (i == 0)
				{
					format = "if ({0}";
					if (openAsElseIf)
						format = "else " + format;
				}
				else
					format = "|| ({0})";
				if (i == (wrappedConditionSegments.Length - 1))
					format += ")";

				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(format, wrappedConditionSegments[i]),
					indentationDepth,
					lineIndex
				));
			}
			return translationResult;
		}

		private TranslationResult SetCaseMatchResultValue(
			TranslationResult translationResult,
			int indentationDepth,
			IEnumerable<TranslatedStatementContentDetails> conditionSegments,
			CSharpName isCaseMatchResultName,
			CSharpName errorRegistrationTokenIfAny,
			int lineIndex)
		{
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth");
			if (conditionSegments == null)
				throw new ArgumentNullException("conditionSegments");
			if (isCaseMatchResultName == null)
				throw new ArgumentNullException("isCaseMatchResultName");
			if (lineIndex < 0)
				throw new ArgumentOutOfRangeException("Must be zero or greater", "lineIndex");

			var conditionSegmentsArray = conditionSegments.ToArray();
			if (conditionSegmentsArray.Length == 0)
				throw new ArgumentException("There must be at least one condition segment");
			if (conditionSegmentsArray.Any(segment => segment == null))
				throw new ArgumentException("Null reference encountered in conditionSegments set");

			for (var i = 0; i < conditionSegmentsArray.Length; i++)
			{
				string format;
				if (i == 0)
					format = "{0} = ({1}";
				else
					format = "|| ({1})";
				if (i == (conditionSegmentsArray.Length - 1))
					format += ");";

				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(format, isCaseMatchResultName.Name, conditionSegmentsArray[i].TranslatedContent),
					(i == 0) ? indentationDepth : indentationDepth + 1,
					lineIndex
				));
			}
			return translationResult;
		}

		private SelectTargetExpressionTranslationData TranslateTargetExpression(Expression targetExpression, TranslationResult translationResult, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (targetExpression == null)
				throw new ArgumentNullException("targetExpression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth");

			// If the target expression is a simple constant then we can skip any work required in evaluating its value
			if (Is<NumericValueToken>(targetExpression) || Is<DateLiteralToken>(targetExpression) || Is<StringToken>(targetExpression) || Is<BuiltInValueToken>(targetExpression))
			{
				return new SelectTargetExpressionTranslationData(
					evaluatedTarget: targetExpression.Tokens.Single(),
					successfullyEvaluatedTargetNameIfRequired: null,
					extendedTranslationResult: translationResult,
					extendedScopeAccessInformation: scopeAccessInformation
				);
			}

			// If it's a single NameToken then we can also avoid some work since we don't need to evaluate it (but do need to apply the usual logic to determine whether it's an undeclared
			// variable). If the target IS a single NameToken, it does not matter at this point whether it is a VBScript value or object type at this point - if it IS an object type then it
			// will need a default parameterless method or property, but that will be called when each comparison is made (this is important; the default member will not be called once and
			// its value stashed, it will be called for EACH comparison - so the NameToken is not wrapped in a VAL call at this point, the must-be-value-type logic will be handled in the EQ
			// implementation in the compat library since EQ only deals with value types, the IS operator deals with object types).
			if (Is<NameToken>(targetExpression))
			{
				var targetNameToken = (NameToken)targetExpression.Tokens.Single();
				if (!scopeAccessInformation.IsDeclaredReference(targetNameToken, _nameRewriter))
				{
					_logger.Warning("Undeclared variable: \"" + targetNameToken.Content + "\" (line " + (targetNameToken.LineIndex + 1) + ")");
					translationResult = translationResult.AddUndeclaredVariables(new[] { targetNameToken });
				}
				return new SelectTargetExpressionTranslationData(
					evaluatedTarget: targetNameToken,
					successfullyEvaluatedTargetNameIfRequired: null,
					extendedTranslationResult: translationResult,
					extendedScopeAccessInformation: scopeAccessInformation
				);
			}

			// If we have to evaluate the target expression then we'll store it in a local variable which will be referenced by the comparisons. This variable needs to be added to the
			// scope access information so that when the comparisons are translated into C#, the code that does so knows that it is a declared local variable and not an undeclared variable
			// that must be accessed through the "Environment References" class. Note that there is currently no way to add a name token to the scope that will not pass through the name
			// rewriter, so the name "target" needs to be one that we do not expect the name rewriter to affect (this is not great since it means that there is an undeclared implicit
			// dependency on the VBScriptNameRewriter implementation but it's all I've got at the moment - since ScopedNameToken is derived from NameToken, I would have to change this
			// around to support a ScopedDoNotRenameNameToken).
			var evaluatedTargetName = _tempNameGenerator(new CSharpName("target"), scopeAccessInformation);
			scopeAccessInformation = scopeAccessInformation.ExtendVariables(new NonNullImmutableList<ScopedNameToken>(new[] {
				new ScopedNameToken(evaluatedTargetName.Name, targetExpression.Tokens.First().LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith)
			}));

			// If error-trapping may be enabled at runtime then we need to wrap the select target evaluation in a HANDLEERROR call. If the evaulation fails then nothing else in the select
			// construct should be considered - so a flag is set to false before evaluation is attempted and set to true if the evaluation was successful, the comparisons work will only
			// be executed if the flag was true. (In some cases, VBScript tries to do *something* if an error occurs but is trapped, but this is not one of them - unlike when the
			// comparisons ARE considered; if any of them raise an error while error-trapping is enabled then they will be considered to match and their child statements will
			// be executed).
			// - This complexity is avoided if error-trapping is definitely not in play
			TranslatedStatementContentDetails evaluatedTargetContent;
			var byRefArgumentIdentifier = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
			var byRefArgumentsToRewrite = byRefArgumentIdentifier.GetByRefArgumentsThatNeedRewriting(
				targetExpression.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
				scopeAccessInformation,
				new NonNullImmutableList<FuncByRefMapping>()
			);
			CSharpName successfullyEvaluatedTargetNameIfRequired;
			if (byRefArgumentsToRewrite.Any())
			{
				// If we're in a function or property and that function / property has by-ref arguments that we then need to pass into further function / property calls
				// in order to evaluate the current conditional, then we need to record those values in temporary references, try to evaulate the value and then push the
				// temporary values back into the original references. This is required in order to be consistent with VBScript and yet also produce code that compiles as
				// C# (which will not let "ref" arguments of the containing function be used in lambdas, which is how we deal with updating by-ref arguments after function
				// or property calls complete).
				scopeAccessInformation = scopeAccessInformation.ExtendVariables(
					byRefArgumentsToRewrite
						.Select(r => new ScopedNameToken(r.To.Name, r.From.LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith))
						.ToNonNullImmutableList()
				);
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(
						"object {0} = null;",
						evaluatedTargetName.Name
					),
					indentationDepth,
					targetExpression.Tokens.First().LineIndex
				));
				if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
					successfullyEvaluatedTargetNameIfRequired = null;
				else
				{
					successfullyEvaluatedTargetNameIfRequired = _tempNameGenerator(new CSharpName("targetWasEvaluated"), scopeAccessInformation);
					translationResult = translationResult.Add(new TranslatedStatement(
						"var " + successfullyEvaluatedTargetNameIfRequired.Name + " = false;",
						indentationDepth,
						targetExpression.Tokens.First().LineIndex
					));
				}
				var byRefMappingOpeningTranslationDetails = byRefArgumentsToRewrite.OpenByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
				translationResult = byRefMappingOpeningTranslationDetails.TranslationResult;
				indentationDepth += byRefMappingOpeningTranslationDetails.DistanceToIndentCodeWithMappedValues;
				if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
				{
					translationResult = translationResult.Add(new TranslatedStatement(
						string.Format(
							"{0}.HANDLEERROR({1}, () => {{",
							_supportRefName.Name,
							scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
						),
						indentationDepth,
						targetExpression.Tokens.First().LineIndex
					));
					indentationDepth++;
				}
				var rewrittenConditionalContent = _statementTranslator.Translate(
					byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(targetExpression, _nameRewriter),
					scopeAccessInformation,
					ExpressionReturnTypeOptions.NotSpecified,
					_logger.Warning
				);
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(
						"{0} = {1};",
						evaluatedTargetName.Name,
						rewrittenConditionalContent.TranslatedContent
					),
					indentationDepth,
					targetExpression.Tokens.First().LineIndex
				));
				if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
				{
					translationResult = translationResult.Add(new TranslatedStatement(
						successfullyEvaluatedTargetNameIfRequired.Name + " = true;",
						indentationDepth,
						targetExpression.Tokens.First().LineIndex
					));
					indentationDepth--;
					translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth, targetExpression.Tokens.First().LineIndex));
				}
				indentationDepth -= byRefMappingOpeningTranslationDetails.DistanceToIndentCodeWithMappedValues;
				translationResult = byRefArgumentsToRewrite.CloseByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
				evaluatedTargetContent = new TranslatedStatementContentDetails(
					evaluatedTargetName.Name,
					rewrittenConditionalContent.VariablesAccessed
				);
			}
			else if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(
						"object {0} = null;",
						evaluatedTargetName.Name
					),
					indentationDepth,
					targetExpression.Tokens.First().LineIndex
				));
				successfullyEvaluatedTargetNameIfRequired = _tempNameGenerator(new CSharpName("targetWasEvaluated"), scopeAccessInformation);
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format("var {0} = false;", successfullyEvaluatedTargetNameIfRequired.Name),
					indentationDepth,
					targetExpression.Tokens.First().LineIndex
				));
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(
						"{0}.HANDLEERROR({1}, () => {{",
						_supportRefName.Name,
						scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
					),
					indentationDepth,
					targetExpression.Tokens.First().LineIndex
				));
				evaluatedTargetContent = _statementTranslator.Translate(targetExpression, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(
						"{0} = {1};",
						evaluatedTargetName.Name,
						evaluatedTargetContent.TranslatedContent
					),
					indentationDepth + 1,
					targetExpression.Tokens.First().LineIndex
				));
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format("{0} = true;", successfullyEvaluatedTargetNameIfRequired.Name),
					indentationDepth + 1,
					targetExpression.Tokens.First().LineIndex
				));
				translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth,targetExpression.Tokens.First().LineIndex));
			}
			else
			{
				evaluatedTargetContent = _statementTranslator.Translate(targetExpression, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
				translationResult = translationResult.Add(new TranslatedStatement(
					string.Format(
						"object {0} = {1};", // Best to declare "object" type rather than "var" in case the SELECT CASE target is Empty (ie. null)
						evaluatedTargetName.Name,
						evaluatedTargetContent.TranslatedContent
					),
					indentationDepth,
					targetExpression.Tokens.First().LineIndex
				));
				successfullyEvaluatedTargetNameIfRequired = null; // Not required since we can don't have to deal with fail cases (since error handling is not enabled)
			}
			var undeclaredVariablesAccessedInTargetExpression = evaluatedTargetContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
			foreach (var undeclaredVariable in undeclaredVariablesAccessedInTargetExpression)
				_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
			translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesAccessedInTargetExpression);

			// Since the target expression wasn't something simple (like a literal or built-in value), then it's stored in a variable so that it's not evaluated for each case
			// comparison (if it's a function call, for example, then the function shouldn't be called for each comparison). However, if it's an object reference, then each
			// comparison is going to try to force it into a value type reference - meaning it will require a parameter-less default method or property. Assuming it has one
			// of these, this WILL be accessed for each comparison (so it's important not to try to force the evaluatedTarget into a VAL at this point - the comparisons
			// will force it into a value type and the getter / default property will be called each time then, as is consistent with VBScript).
			return new SelectTargetExpressionTranslationData(
				evaluatedTarget: new NameToken(evaluatedTargetName.Name, targetExpression.Tokens.First().LineIndex),
				successfullyEvaluatedTargetNameIfRequired: successfullyEvaluatedTargetNameIfRequired,
				extendedTranslationResult: translationResult,
				extendedScopeAccessInformation: scopeAccessInformation
			);
		}

		private TranslatedStatementContentDetails GetComparison(
			IToken evaluatedTarget,
			Expression value,
			bool isFirstValueInCaseSet,
			ScopeAccessInformation scopeAccessInformation)
		{
			if (evaluatedTarget == null)
				throw new ArgumentNullException("evaluatedTargetName");
			if (value == null)
				throw new ArgumentNullException("value");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var translatedTargetToCompareTo = _statementTranslator.Translate(
				new Expression(new[] { evaluatedTarget }),
				scopeAccessInformation,
				ExpressionReturnTypeOptions.NotSpecified,
				_logger.Warning
			);

			// If the target case is a numeric literal then the first option in each case set must be parseable as a number
			// If the target case is a numeric literal the non-first options in each case set need not be parseable as numbers but flexible matching will be applied (1 and "1" are considered equal)
			// - Before dealing with these, if the current value is a numeric constant and the target case is a numeric literal then we can do a straight EQ call on them
			// Note: There is no need to specify ExpressionReturnTypeOptions.Value here, it just adds noise to the output since the EQ method will ensure that the left and right values are both
			// value types (since EQ only compares value types - unlike IS, which only compares object types).
			// TODO: Need to do the same for date literals too?
			var evaluatedExpression = _statementTranslator.Translate(value, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
			if ((evaluatedTarget is NumericValueToken) && IsNumericLiteral((NumericValueToken)evaluatedTarget))
			{
				if (Is<NumericValueToken>(value))
				{
					return new TranslatedStatementContentDetails(
						string.Format(
							"{0}.EQ({1}, {2})",
							_supportRefName.Name,
							translatedTargetToCompareTo.TranslatedContent,
							evaluatedExpression.TranslatedContent
						),
						evaluatedExpression.VariablesAccessed
					);
				}

				if (isFirstValueInCaseSet)
				{
					return new TranslatedStatementContentDetails(
						string.Format(
							"{0}.EQ({1}, {0}.NUM({2}))",
							_supportRefName.Name,
							translatedTargetToCompareTo.TranslatedContent,
							evaluatedExpression.TranslatedContent
						),
						evaluatedExpression.VariablesAccessed
					);
				}

				return new TranslatedStatementContentDetails(
					string.Format(
						"{0}.EQish({1}, {2})",
						_supportRefName.Name,
						translatedTargetToCompareTo.TranslatedContent,
						evaluatedExpression.TranslatedContent
					),
					evaluatedExpression.VariablesAccessed
				);
			}

			// If the case value is a numeric literal then the target must be parseable as a number - TODO: Need to do the same for date literals too?
			if (IsNumericLiteral(value))
			{
				if (evaluatedTarget is NumericValueToken)
				{
					return new TranslatedStatementContentDetails(
						string.Format(
							"{0}.EQ({1}, {2})",
							_supportRefName.Name,
							translatedTargetToCompareTo.TranslatedContent,
							evaluatedExpression.TranslatedContent
						),
						evaluatedExpression.VariablesAccessed
					);
				}

				return new TranslatedStatementContentDetails(
					string.Format(
						"{0}.EQ({0}.NUM({1}), {2})",
						_supportRefName.Name,
						translatedTargetToCompareTo.TranslatedContent,
						evaluatedExpression.TranslatedContent
					),
					evaluatedExpression.VariablesAccessed
				);
			}

			// If neither value (target nor case option) are numeric literals, then no flexible matching is applied (there is apparently no special behaviour applied to string literals in either the target
			// expression nor any value within a case set)
			return new TranslatedStatementContentDetails(
				string.Format(
					"{0}.EQ({1}, {2})",
					_supportRefName.Name,
					translatedTargetToCompareTo.TranslatedContent,
					evaluatedExpression.TranslatedContent
				),
				evaluatedExpression.VariablesAccessed
			);
		}

		/// <summary>
		/// VBScript does not consider -1 to be a numeric literal, it is a subtraction operation against the numeric literal 1 (so special rules around numeric literals do not apply to negative values)
		/// </summary>
		private bool IsNumericLiteral(Expression expression)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");

			return Is<NumericValueToken>(expression) && IsNumericLiteral((NumericValueToken)expression.Tokens.Single());
		}

		/// <summary>
		/// VBScript does not consider -1 to be a numeric literal, it is a subtraction operation against the numeric literal 1 (so special rules around numeric literals do not apply to negative values)
		/// </summary>
		private bool IsNumericLiteral(NumericValueToken numericValueToken)
		{
			if (numericValueToken == null)
				throw new ArgumentNullException("numericValueToken");

			return !numericValueToken.Content.StartsWith("-");
		}

		private bool Is<TSingleTokenType>(Expression expression) where TSingleTokenType : IToken
		{
			if (expression == null)
				throw new ArgumentNullException("expression");

			return (expression.Tokens.Count() == 1) && (expression.Tokens.Single() is TSingleTokenType);
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

		private class SelectTargetExpressionTranslationData
		{
			public SelectTargetExpressionTranslationData(IToken evaluatedTarget, CSharpName successfullyEvaluatedTargetNameIfRequired, TranslationResult extendedTranslationResult, ScopeAccessInformation extendedScopeAccessInformation)
			{
				if (evaluatedTarget == null)
					throw new ArgumentNullException("evaluatedTarget");
				if (extendedTranslationResult == null)
					throw new ArgumentNullException("extendedTranslationResult");
				if (extendedScopeAccessInformation == null)
					throw new ArgumentNullException("extendedScopeAccessInformation");

				EvaluatedTarget = evaluatedTarget;
				SuccessfullyEvaluatedTargetNameIfRequired = successfullyEvaluatedTargetNameIfRequired;
				ExtendedTranslationResult = extendedTranslationResult;
				ExtendedScopeAccessInformation = extendedScopeAccessInformation;
			}

			/// <summary>
			/// This will always be a single token since either the target expression is a simple constant or a single NameToken (in which case the token for it will be present here) or
			/// it's something that needs evaluating, in which case a NameToken will be generated for the temporary value that the expression result will be set into
			/// </summary>
			public IToken EvaluatedTarget { get; private set; }

			/// <summary>
			/// If the target expresssion is simple and does not require evaluation, or if there is no error-handling (in which case a failed evaluation would result in a termination of
			/// work at this point), then this will be null.
			/// </summary>
			public CSharpName SuccessfullyEvaluatedTargetNameIfRequired { get; private set; }

			/// <summary>
			/// This will never be null
			/// </summary>
			public TranslationResult ExtendedTranslationResult { get; private set; }

			/// <summary>
			/// This will never be null
			/// </summary>
			public ScopeAccessInformation ExtendedScopeAccessInformation { get; private set; }
		}

		private class ConditionMatchingByRefArgAliasingDetails
		{
			public ConditionMatchingByRefArgAliasingDetails(
				NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite,
				FuncByRefMappingList_Extensions.ByRefReplacementTranslationResultDetails byRefAliasWrappingDetails,
				CSharpName caseValueMatchResultName)
			{
				if (byRefArgumentsToRewrite == null)
					throw new ArgumentNullException("byRefArgumentsToRewrite");
				if (!byRefArgumentsToRewrite.Any())
					throw new ArgumentException("byRefArgumentsToRewrite may not be an empty set, otherwise there's no point to creating this instance");
				if (byRefAliasWrappingDetails == null)
					throw new ArgumentNullException("byRefAliasWrappingDetails");
				if (caseValueMatchResultName == null)
					throw new ArgumentNullException("caseValueMatchResultName");

				ByRefArgumentsToRewrite = byRefArgumentsToRewrite;
				ByRefAliasWrappingDetails = byRefAliasWrappingDetails;
				CaseValueMatchResultName = caseValueMatchResultName;
			}

			/// <summary>
			/// This will never be null or empty
			/// </summary>
			public NonNullImmutableList<FuncByRefMapping> ByRefArgumentsToRewrite { get; private set; }

			/// <summary>
			/// This will never be null
			/// </summary>
			public FuncByRefMappingList_Extensions.ByRefReplacementTranslationResultDetails ByRefAliasWrappingDetails { get; private set; }

			/// <summary>
			/// This will never be null
			/// </summary>
			public CSharpName CaseValueMatchResultName { get; private set; }
		}
	}
}
