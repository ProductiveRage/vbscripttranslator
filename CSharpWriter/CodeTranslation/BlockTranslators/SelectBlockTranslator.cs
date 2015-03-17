using System;
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

            // 1. Case values are lazily evaluated; as soon as one is matched, no more are considered.
            // 2. There is no fall-through from one section to another; the first matched section (if any) is processed and no more are considered.
            
            // TODO: Special handling for zero cases - the target should still be evaluated (document this)

            // TODO: Error-handling

            // TODO: ByRef function argument rewriting (in conjunction with error-handling on on its own)

            var translationResult = TranslationResult.Empty;
            foreach (var openingComment in selectBlock.OpeningComments)
                translationResult = base.TryToTranslateComment(translationResult, openingComment, scopeAccessInformation, indentationDepth);

            // Note: Don't try to do anything clever with evaulated-target, like avoid declaring a variable if it's just a number or string or variable reference since that means having
            // to consider too many possibilities here - eg. Is the type of the number important? Does the variable reference need passing through a VAL call? Much better to just evaluate
            // it using the standard mechanisms and stash it in a temporary variable that case options can be compared against.
            var evaluatedTargetName = _tempNameGenerator(new CSharpName("selectCase"), scopeAccessInformation);
            scopeAccessInformation = scopeAccessInformation.ExtendVariables(new NonNullImmutableList<ScopedNameToken>(new[] {
                new ScopedNameToken(evaluatedTargetName.Name, selectBlock.Expression.Tokens.First().LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith)
            }));
            var evaluatedTargetContent = _statementTranslator.Translate(selectBlock.Expression, scopeAccessInformation, ExpressionReturnTypeOptions.Value, _logger.Warning);
            foreach (var undeclaredVariable in evaluatedTargetContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter))
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
            translationResult = translationResult.Add(new TranslatedStatement(
                string.Format(
                    "object {0} = {1};", // Best to declare "object" type rather than "var" in case the SELECT CASE target is Empty (ie. null)
                    evaluatedTargetName.Name,
                    evaluatedTargetContent.TranslatedContent
                ),
                indentationDepth
            ));

            foreach (var indexedCaseBlock in selectBlock.Content.Select((c, i) => new { Index = i, CaseBlock = c }))
            {
                var explicitOptionsCaseBlock = indexedCaseBlock.CaseBlock as SelectBlock.CaseBlockExpressionSegment;
                if (explicitOptionsCaseBlock != null)
                {
                    var targetAsNumericValueTokenIfApplicable = Is<NumericValueToken>(selectBlock.Expression)
                        ? (NumericValueToken)selectBlock.Expression.Tokens.Single()
                        : null;
                    var conditions = explicitOptionsCaseBlock.Values
                        .Select((value, index) => GetComparison(
                            evaluatedTargetName,
                            targetAsNumericValueTokenIfApplicable,
                            value,
                            isFirstValueInCaseSet: (index == 0),
                            scopeAccessInformation: scopeAccessInformation
                        ))
                        .ToArray(); // Call ToArray since we're always going to enumerate the collection twice below, no point doing the work twice
                    foreach (var undeclaredVariable in conditions.SelectMany(c => c.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)))
                        _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                    var combinedCondition = (explicitOptionsCaseBlock.Values.Count() == 1)
                        ? conditions.Single().TranslatedContent
                        : ("(" + string.Join(") || (", conditions.Select(c => c.TranslatedContent)) + ")");

                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "{0} ({1})",
                            (indexedCaseBlock.Index == 0) ? "if" : "else if",
                            combinedCondition
                        ),
                        indentationDepth
                    ));
                    translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                    translationResult = translationResult.Add(
                        Translate(indexedCaseBlock.CaseBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, indentationDepth + 1)
                    );
                    translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
                }
                else
                {
                    var defaultCaseIsTheOnlyCase = (selectBlock.Content.Count() == 1);
                    if (!defaultCaseIsTheOnlyCase)
                    {
                        translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth));
                        translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                        indentationDepth++;
                    }
                    translationResult = translationResult.Add(
                        Translate(indexedCaseBlock.CaseBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, indentationDepth)
                    );
                    if (!defaultCaseIsTheOnlyCase)
                    {
                        indentationDepth--;
                        translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
                    }
                }
            }

            return translationResult;
		}

        private TranslatedStatementContentDetails GetComparison(
            CSharpName evaluatedTargetName,
            NumericValueToken targetAsNumericValueTokenIfApplicable,
            Expression value,
            bool isFirstValueInCaseSet,
            ScopeAccessInformation scopeAccessInformation)
        {
            if (evaluatedTargetName == null)
                throw new ArgumentNullException("evaluatedTargetName");
            if (value == null)
                throw new ArgumentNullException("value");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            // If the target case is a numeric literal then the first option in each case set must be parseable as a number
            // If the target case is a numeric literal the non-first options in each case set need not be parseable as numbers but flexible matching will be applied (1 and "1" are considered equal)
            // - Before dealing with these, if the current value is a numeric constant and the target case is a numeric literal then we can do a straight EQ call on them
            var evaluatedExpression = _statementTranslator.Translate(value, scopeAccessInformation, ExpressionReturnTypeOptions.Value, _logger.Warning);
            if ((targetAsNumericValueTokenIfApplicable != null) && IsNumericLiteral(targetAsNumericValueTokenIfApplicable))
            {
                if (Is<NumericValueToken>(value))
                {
                    return new TranslatedStatementContentDetails(
                        string.Format(
                            "{0}.EQ({1}, {2})",
                            _supportRefName.Name,
                            evaluatedTargetName.Name,
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
                            evaluatedTargetName.Name,
                            evaluatedExpression.TranslatedContent
                        ),
                        evaluatedExpression.VariablesAccessed
                    );
                }
                
                return new TranslatedStatementContentDetails(
                    string.Format(
                        "{0}.EQish({1}, {2})",
                        _supportRefName.Name,
                        evaluatedTargetName.Name,
                        evaluatedExpression.TranslatedContent
                    ),
                    evaluatedExpression.VariablesAccessed
                );
            }

            // If the case value is a numeric literal then the target must be parseable as a number
            if (IsNumericLiteral(value))
            {
                if (targetAsNumericValueTokenIfApplicable != null)
                {
                    return new TranslatedStatementContentDetails(
                        string.Format(
                            "{0}.EQ({1}, {2})",
                            _supportRefName.Name,
                            evaluatedTargetName.Name,
                            evaluatedExpression.TranslatedContent
                        ),
                        evaluatedExpression.VariablesAccessed
                    );
                }

                return new TranslatedStatementContentDetails(
                    string.Format(
                        "{0}.EQ({1}, {0}.NUM({2}))",
                        _supportRefName.Name,
                        evaluatedTargetName.Name,
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
                    evaluatedTargetName.Name,
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
    }
}
