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

            // Notes:
            // 1. Case values are lazily evaluated; as soon as one is matched, no more are considered.
            // 2. There is no fall-through from one section to another; the first matched section (if any) is processed and no more are considered.
            
            var translationResult = TranslationResult.Empty;
            foreach (var openingComment in selectBlock.OpeningComments)
                translationResult = base.TryToTranslateComment(translationResult, openingComment, scopeAccessInformation, indentationDepth);

            // If the target expression is a simple constant then we can skip any work required in evaluating its value. If it's a single NameToken then we can also avoid some work since we
            // don't need to evaluate it (but do need to apply the usual logic to determine whether it's an undeclared variable). If the target IS a single NameToken, it does not matter at this
            // point whether it is a VBScript value or object type at this point - if it IS an object type then it will need a default parameterless method or property, but that will be called
            // when each comparison is made (this is important; the default member will not be called once and its value stashed, it will be called for EACH comparison - so the NameToken is
            // not wrapped in a VAL call at this point, the must-be-value-type logic will be handled in the EQ implementation in the compat library since EQ only deals with value types, the
            // IS operator deals with object types).
            CSharpName successfullyEvaluatedTargetNameIfRequired;
            IToken evaluatedTarget;
            if (Is<NumericValueToken>(selectBlock.Expression) || Is<DateLiteralToken>(selectBlock.Expression) || Is<StringToken>(selectBlock.Expression) || Is<BuiltInValueToken>(selectBlock.Expression))
            {
                evaluatedTarget = selectBlock.Expression.Tokens.Single();
                successfullyEvaluatedTargetNameIfRequired = null;
            }
            else if (Is<NameToken>(selectBlock.Expression))
            {
                var targetNameToken = (NameToken)selectBlock.Expression.Tokens.Single();
                if (!scopeAccessInformation.IsDeclaredReference(_nameRewriter.GetMemberAccessTokenName(targetNameToken), _nameRewriter))
                {
                    _logger.Warning("Undeclared variable: \"" + targetNameToken.Content + "\" (line " + (targetNameToken.LineIndex + 1) + ")");
                    translationResult = translationResult.AddUndeclaredVariables(new[] { targetNameToken });
                }
                evaluatedTarget = targetNameToken;
                successfullyEvaluatedTargetNameIfRequired = null;
            }
            else
            {
                // If we have to evaluate the target expression then we'll store it in a local variable which will be referenced by the comparisons. This variable needs to be added to the
                // scope access information so that when the comparisons are translated into C#, the code that does so knows that it is a declared local variable and not an undeclared variable
                // that must be accessed through the "Environment References" class. Note that there is currently no way to add a name token to the scope that will not pass through the name
                // rewriter, so the name "target" needs to be one that we do not expect the name rewriter to affect (this is not great since it means that there is an undeclared implicit
                // dependency on the VBScriptNameRewriter implementation but it's all I've got at the moment - since ScopedNameToken is derived from NameToken, I would have to change this
                // around to support a ScopedDoNotRenameNameToken).
                var evaluatedTargetName = _tempNameGenerator(new CSharpName("target"), scopeAccessInformation);
                scopeAccessInformation = scopeAccessInformation.ExtendVariables(new NonNullImmutableList<ScopedNameToken>(new[] {
                    new ScopedNameToken(evaluatedTargetName.Name, selectBlock.Expression.Tokens.First().LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith)
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
                    selectBlock.Expression.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
                    scopeAccessInformation,
                    new NonNullImmutableList<FuncByRefMapping>()
                );
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
                        indentationDepth
                    ));
                    if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
                        successfullyEvaluatedTargetNameIfRequired = null;
                    else
                    {
                        successfullyEvaluatedTargetNameIfRequired = _tempNameGenerator(new CSharpName("targetWasEvaluated"), scopeAccessInformation);
                        translationResult = translationResult.Add(new TranslatedStatement(
                            "var " + successfullyEvaluatedTargetNameIfRequired.Name + " = false;",
                            indentationDepth
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
                            indentationDepth
                        ));
                        indentationDepth++;
                    }
                    var rewrittenConditionalContent = _statementTranslator.Translate(
                        byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(selectBlock.Expression, _nameRewriter),
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
                        indentationDepth
                    ));
                    if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
                    {
                        translationResult = translationResult.Add(new TranslatedStatement(
                            successfullyEvaluatedTargetNameIfRequired.Name + " = true;",
                            indentationDepth
                        ));
                        indentationDepth--;
                        translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth));
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
                        indentationDepth
                    ));
                    successfullyEvaluatedTargetNameIfRequired = _tempNameGenerator(new CSharpName("targetWasEvaluated"), scopeAccessInformation);
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format("var {0} = false;", successfullyEvaluatedTargetNameIfRequired.Name),
                        indentationDepth
                    ));
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "{0}.HANDLEERROR({1}, () => {{",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ));
                    evaluatedTargetContent = _statementTranslator.Translate(selectBlock.Expression, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "{0} = {1};",
                            evaluatedTargetName.Name,
                            evaluatedTargetContent.TranslatedContent
                        ),
                        indentationDepth + 1
                    ));
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format("{0} = true;", successfullyEvaluatedTargetNameIfRequired.Name),
                        indentationDepth + 1
                    ));
                    translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth));
                }
                else
                {
                    evaluatedTargetContent = _statementTranslator.Translate(selectBlock.Expression, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning);
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "object {0} = {1};", // Best to declare "object" type rather than "var" in case the SELECT CASE target is Empty (ie. null)
                            evaluatedTargetName.Name,
                            evaluatedTargetContent.TranslatedContent
                        ),
                        indentationDepth
                    ));
                    successfullyEvaluatedTargetNameIfRequired = null; // Not required since we can don't have to deal with fail cases (since error handling is not enabled)
                }
                var undeclaredVariablesAccessedInTargetExpression = evaluatedTargetContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
                foreach (var undeclaredVariable in undeclaredVariablesAccessedInTargetExpression)
                    _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesAccessedInTargetExpression);

                // If the expression was not evaluated then no comparisons work is required - note that this check is not required if there are zero comparisons to deal with. In
                // this case, VBScript would still evaluate the target even though there's nothing to compare it to (and we remain consistent with that; the target evaluation
                // occurs above but we do nothing more if there are no comparison values).
                if ((successfullyEvaluatedTargetNameIfRequired != null) && selectBlock.Content.Any())
                {
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format("if ({0})", successfullyEvaluatedTargetNameIfRequired.Name),
                        indentationDepth
                    ));
                    translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                    indentationDepth++;
                }

                // Since the target expression wasn't something simple (like a literal or built-in value), then it's stored in a variable so that it's not evaluated for each case
                // comparison (if it's a function call, for example, then the function shouldn't be called for each comparison). However, if it's an object reference, then each
                // comparison is going to try to force it into a value type reference - meaning it will require a parameter-less default method or property. Assuming it has one
                // of these, this WILL be accessed for each comparison (so it's important not to try to force the evaluatedTarget into a VAL at this point - the comparisons
                // will force it into a value type and the getter / default property will be called each time then, as is consistent with VBScript).
                evaluatedTarget = new NameToken(evaluatedTargetName.Name, selectBlock.Expression.Tokens.First().LineIndex);
            }

            var numberOfIndentsRequiredForErrorTrappingStructure = 0;
            var annotatedCaseBlocks = selectBlock.Content.Select((c, i) => {
                var lastIndex = selectBlock.Content.Count() - 1;
                return new
                {
                    IsFirstCaseBlock = (i == 0),
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
                    var conditions = explicitOptionsCaseBlock.Values
                        .Select((value, index) => GetComparison(
                            evaluatedTarget,
                            value,
                            isFirstValueInCaseSet: (index == 0),
                            scopeAccessInformation: scopeAccessInformation
                        ))
                        .ToArray(); // Call ToArray since we're always going to enumerate the collection twice below, no point doing the work twice
                    var undeclaredVariablesInCondition = conditions.SelectMany(c => c.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter));
                    foreach (var undeclaredVariable in undeclaredVariablesInCondition)
                        _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                    translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesInCondition);
                    
                    // If error-trapping is not enabled then generate an "if" or "else if" conditional that checks each value using an or, so that if one matches then any subsequent values are not evaluated
                    if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
                    {
                        conditions = conditions
                            .Select(condition => new TranslatedStatementContentDetails(
                                string.Format(
                                    "{0}.IF({1})",
                                    _supportRefName.Name,
                                    condition.TranslatedContent
                                ),
                                condition.VariablesAccessed
                            ))
                            .ToArray();
                        var combinedCondition = (explicitOptionsCaseBlock.Values.Count() == 1)
                            ? conditions.Single().TranslatedContent
                            : ("(" + string.Join(") || (", conditions.Select(c => c.TranslatedContent)) + ")");
                        translationResult = translationResult.Add(new TranslatedStatement(
                            string.Format(
                                "{0} ({1})",
                                annotatedCaseBlock.IsFirstCaseBlock ? "if" : "else if",
                                combinedCondition
                            ),
                            indentationDepth
                        ));
                    }
                    else
                    {
                        // If error-trapping IS potentially enabled, then these conditions must each be wrapped in a call to the IF support function that takes the error token and deals with evaluating
                        // the value; if the evaluation causes an error and error-trapping is enabled at runtime then VBScript would "resume" into the next statement, meaning the block within the CASE
                        // (and so the IF support function knows to return true in this case).
                        // - Since this involves more code for each value comparison, if the CASE has multiple values then they will be split onto multiple lines in the emitted C# code
                        // - Also note that if error-trapping is enabled then we need to change the structure slightly, from "if.. elseif .. else if.. else" to using nested "if.. else { if.. else { } } }"
                        //   since each condition will be evaluated within a lambda and so may need rewriting if the expression references any ByRef arguments of the containing function (where applicable)
                        //   because "ref" references may not be accessed within lambdas in C#. TODO: I haven't dealt with this rewriting yet, but I've laid out this structure here so that I'm able to do
                        //   that work soon (2015-05-28: The potential rewriting of the target expression has been dealt with, but not any the case values)
                        conditions = conditions
                            .Select(condition => new TranslatedStatementContentDetails(
                                string.Format(
                                    "{0}.IF(() => {1}, {2})",
                                    _supportRefName.Name,
                                    condition.TranslatedContent,
                                    scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                                ),
                                condition.VariablesAccessed
                            ))
                            .ToArray();
                        if (conditions.Count() == 1)
                        {
                            translationResult = translationResult.Add(new TranslatedStatement(
                                string.Format("if ({0})", conditions.Single().TranslatedContent),
                                indentationDepth
                            ));
                        }
                        else
                        {
                            translationResult = translationResult.Add(new TranslatedStatement(
                                string.Format("if (({0})", conditions.First().TranslatedContent),
                                indentationDepth
                            ));
                            foreach (var condition in conditions.Skip(1).Reverse().Skip(1).Reverse()) // Every condition except the first and last
                            {
                                translationResult = translationResult.Add(new TranslatedStatement(
                                    string.Format("|| ({0})", condition.TranslatedContent),
                                    indentationDepth
                                ));
                            }
                            translationResult = translationResult.Add(new TranslatedStatement(
                                string.Format("|| ({0}))", conditions.Last().TranslatedContent),
                                indentationDepth
                            ));
                        }
                    }

                    translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                    translationResult = translationResult.Add(
                        Translate(
                            annotatedCaseBlock.CaseBlock.Statements.ToNonNullImmutableList(),
                            scopeAccessInformation.SetParent(selectBlock),
                            indentationDepth + 1
                        )
                    );
                    translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));

                    // See the note above about potentially having to rewrite conditions that include "ref" references - that requires that we change the structure from a more natual "if.. else if.. else if.."
                    // arrangement to having to nest them (eg. "if { .. else { if.. else { } } }")
                    if ((scopeAccessInformation.ErrorRegistrationTokenIfAny != null) && !annotatedCaseBlock.IsCaseLastBlockWithValuesToCheck)
                    {
                        translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth));
                        translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                        indentationDepth++;
                        numberOfIndentsRequiredForErrorTrappingStructure++;
                    }
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
                        Translate(
                            annotatedCaseBlock.CaseBlock.Statements.ToNonNullImmutableList(),
                            scopeAccessInformation.SetParent(selectBlock),
                            indentationDepth
                        )
                    );
                    if (!defaultCaseIsTheOnlyCase)
                    {
                        indentationDepth--;
                        translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
                    }
                }
            }
            for (var i = 0; i < numberOfIndentsRequiredForErrorTrappingStructure; i++)
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }

            // If error-trapping may be active at runtime then the meat of translated content will have been wrapped in an "if", based upon whether the select target was successfully evaluated (in which case
            // we'll need to close that content here). Note that this will not have been the case if there were no expression to compare the target to (VBScript allows this - it evaluates the target expression
            // but then does nothing more).
            if ((successfullyEvaluatedTargetNameIfRequired != null) && selectBlock.Content.Any())
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }

            return translationResult;
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
            var evaluatedExpression = _statementTranslator.Translate(value, scopeAccessInformation, ExpressionReturnTypeOptions.Value, _logger.Warning);
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

            // If the case value is a numeric literal then the target must be parseable as a number
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
    }
}
