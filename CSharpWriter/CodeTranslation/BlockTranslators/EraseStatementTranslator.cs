using System;
using System.Linq;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class EraseStatementTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public EraseStatementTranslator(
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

        public TranslationResult Translate(EraseStatement eraseStatement, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (eraseStatement == null)
                throw new ArgumentNullException("eraseStatement");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // We need to work out what tokens in the target(s) or their argument(s) reference any by-ref arguments in the containing function (where applicable). For the case
            // where there is a single target which is itself a single NameToken, if this is a by-ref argument of the containing function then this will definitely need rewriting.
            // The FuncByRefArgumentMapper doesn't know about this since it is not aware that that token will be reference inside a "targetSetter" lambda, so this needs to be
            // checked explicitly first. In an ideal world, this would be done later on, with the other "success case" logic - but then we'd also have to replicate the functionality
            // in the FuncByRefArgumentMapper that removes any duplicates from the byRefArgumentsToRewrite set (preferring any with read-write mappings over read-only). After this
            // case is handled, all other targets (and any arguments they have) can be rewritten. These won't introduce any "surprising" lambdas, but if there is any error-trapping
            // involved then the erase target evaluation will be wrapped in a HANDLEERROR lambda, but the FuncByRefArgumentMapper IS aware of that and will deal with it accordingly.
            var byRefMapper = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
            var byRefArgumentsToRewrite = new NonNullImmutableList<FuncByRefMapping>();
            if (eraseStatement.Targets.Count() == 1)
            {
                var singleEraseTargetForByRefAliasingConsideration = eraseStatement.Targets.Single();
                if ((singleEraseTargetForByRefAliasingConsideration.Target.Tokens.Count() == 1)
                && (singleEraseTargetForByRefAliasingConsideration.ArgumentsIfAny == null)
                && !singleEraseTargetForByRefAliasingConsideration.WrappedInBraces)
                {
                    var singleTargetNameToken = singleEraseTargetForByRefAliasingConsideration.Target.Tokens.Single() as NameToken;
                    var containingFunctionOrProperty = scopeAccessInformation.ScopeDefiningParent as AbstractFunctionBlock;
                    if ((singleTargetNameToken != null) && (containingFunctionOrProperty != null))
                    {
                        var targetByRefFunctionArgumentIfApplicable = containingFunctionOrProperty.Parameters
                            .Where(p => p.ByRef)
                            .FirstOrDefault(p => _nameRewriter.GetMemberAccessTokenName(p.Name) == _nameRewriter.GetMemberAccessTokenName(singleTargetNameToken));
                        if (targetByRefFunctionArgumentIfApplicable != null)
                        {
                            byRefArgumentsToRewrite = byRefArgumentsToRewrite.Add(new FuncByRefMapping(
                                targetByRefFunctionArgumentIfApplicable.Name,
                                _tempNameGenerator(new CSharpName("byrefalias"), scopeAccessInformation),
                                mappedValueIsReadOnly: false
                            ));
                        }
                    }
                }
            }
            foreach (var targetDetails in eraseStatement.Targets)
            {
                byRefArgumentsToRewrite = byRefMapper.GetByRefArgumentsThatNeedRewriting(
                    targetDetails.Target.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
                    scopeAccessInformation,
                    byRefArgumentsToRewrite
                );
                if (targetDetails.ArgumentsIfAny != null)
                {
                    foreach (var argument in targetDetails.ArgumentsIfAny)
                    {
                        byRefArgumentsToRewrite = byRefMapper.GetByRefArgumentsThatNeedRewriting(
                            argument.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
                            scopeAccessInformation,
                            byRefArgumentsToRewrite
                        );
                    }
                }
            }
            if (byRefArgumentsToRewrite.Any())
            {
                eraseStatement = new EraseStatement(
                    eraseStatement.Targets.Select(targetDetails =>
                    {
                        var rewrittenTarget = byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(targetDetails.Target, _nameRewriter);
                        if (targetDetails.ArgumentsIfAny == null)
                            return new EraseStatement.TargetDetails(rewrittenTarget, null, targetDetails.WrappedInBraces);
                        return new EraseStatement.TargetDetails(
                            rewrittenTarget,
                            targetDetails.ArgumentsIfAny.Select(argument => byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(argument, _nameRewriter)),
                            targetDetails.WrappedInBraces
                        );
                    }),
                    eraseStatement.KeywordLineIndex
                );
            }

            // If the ERASE call is invalid (eg. zero targets "ERASE" or multiple "ERASE a, b" or not a possible by-ref target "ERASE (a)" or "ERASE a.Name" or an invalid array /
            // reference / method call "ERASE a()") then evaluate the targets (to be consistent with VBScript's behaviour) but then raise an error.
            string exceptionStatementIfTargetConfigurationIsInvalid;
            if (eraseStatement.Targets.Count() != 1)
            {
                exceptionStatementIfTargetConfigurationIsInvalid = string.Format(
                    "throw new Exception(\"Wrong number of arguments: 'Erase' (line {0})\");",
                    eraseStatement.KeywordLineIndex + 1
                );
            }
            else
            {
                var eraseTargetToValidate = eraseStatement.Targets.Single();
                if ((eraseTargetToValidate.WrappedInBraces) || (eraseTargetToValidate.Target.Tokens.Count() > 1) || !(eraseTargetToValidate.Target.Tokens.Single() is NameToken))
                {
                    // "Erase (a)" is invalid, it would result in "a" being passed by-val, which would be senseless when trying to erase a dynamic array
                    // "Erase a.Roles" is invalid, the target must be a direct reference (again, since an indirect reference like this would not be passed by-ref)
                    exceptionStatementIfTargetConfigurationIsInvalid = string.Format(
                        "throw new TypeMismatchException(\"'Erase' (line {0})\");",
                        eraseStatement.KeywordLineIndex + 1
                    );
                }
                else
                {
                    // Ensure that the single NameToken in the single erase target is a variable (a function call will result in a "Type mismatch" error)
                    var singleTargetNameToken = (NameToken)eraseTargetToValidate.Target.Tokens.Single();
                    var targetReferenceDetails = scopeAccessInformation.TryToGetDeclaredReferenceDetails(_nameRewriter.GetMemberAccessTokenName(singleTargetNameToken), _nameRewriter);
                    if ((targetReferenceDetails != null) && (targetReferenceDetails.ReferenceType != ReferenceTypeOptions.Variable))
                    {
                        // Note: If the variable has not been declared then targetReferenceDetails will be null, but that means that it will become an undeclared variable later on,
                        // it means that it's definitely not a function
                        exceptionStatementIfTargetConfigurationIsInvalid = string.Format(
                            "throw new TypeMismatchException(\"'Erase' (line {0})\");",
                            eraseStatement.KeywordLineIndex + 1
                        );
                    }
                    else
                        exceptionStatementIfTargetConfigurationIsInvalid = null;
                }
            }

            var translationResult = TranslationResult.Empty;
            int numberOfIndentationLevelsToWithDrawAfterByRefArgumentsProcessed;
            if (byRefArgumentsToRewrite.Any())
            {
                scopeAccessInformation = scopeAccessInformation.ExtendVariables(
                    byRefArgumentsToRewrite
                        .Select(r => new ScopedNameToken(r.To.Name, r.From.LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith))
                        .ToNonNullImmutableList()
                );
                var byRefMappingOpeningTranslationDetails = byRefArgumentsToRewrite.OpenByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
                translationResult = byRefMappingOpeningTranslationDetails.TranslationResult;
                numberOfIndentationLevelsToWithDrawAfterByRefArgumentsProcessed = byRefMappingOpeningTranslationDetails.DistanceToIndentCodeWithMappedValues;
                indentationDepth += numberOfIndentationLevelsToWithDrawAfterByRefArgumentsProcessed;
            }
            else
                numberOfIndentationLevelsToWithDrawAfterByRefArgumentsProcessed = 0;
            if (exceptionStatementIfTargetConfigurationIsInvalid != null)
            {
                foreach (var target in eraseStatement.Targets)
                {
                    var targetExpressionTokens = target.Target.Tokens.ToList();
                    if (target.ArgumentsIfAny != null)
                    {
                        targetExpressionTokens.Add(new OpenBrace(targetExpressionTokens.Last().LineIndex));
                        foreach (var indexedArgument in target.ArgumentsIfAny.Select((a, i) => new { Index = i, Argument = a }))
                        {
                            if (indexedArgument.Index > 0)
                                targetExpressionTokens.Add(new ArgumentSeparatorToken(targetExpressionTokens.Last().LineIndex));
                            targetExpressionTokens.AddRange(indexedArgument.Argument.Tokens);
                        }
                        targetExpressionTokens.Add(new CloseBrace(targetExpressionTokens.Last().LineIndex));
                    }
                    var translatedTarget = _statementTranslator.Translate(
                        new Expression(targetExpressionTokens),
                        scopeAccessInformation,
                        ExpressionReturnTypeOptions.NotSpecified,
                        _logger.Warning
                    );
                    var undeclaredVariablesReferencedByTarget = translatedTarget.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
                    foreach (var undeclaredVariable in undeclaredVariablesReferencedByTarget)
                        _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "var {0} = {1};",
                            _tempNameGenerator(new CSharpName("invalidEraseTarget"), scopeAccessInformation).Name,
                            translatedTarget.TranslatedContent
                        ),
                        indentationDepth
                    ));
                    translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesReferencedByTarget);
                }
                translationResult = translationResult.Add(new TranslatedStatement(exceptionStatementIfTargetConfigurationIsInvalid, indentationDepth));
            }
            else
            {
                // If there are no target arguments then we use the ERASE signature that takes only the target (by-ref). Otherwise call the signature that tries to map the
                // arguments as indices on an array and then erases that element (which must also be an array) - in this case the target need not be passed by-ref.
                // - We know that the ArgumentsIfAny set will be null if there are no items in it, since a non-null-but-empty set is an error condition handled above
                var singleEraseTarget = eraseStatement.Targets.Single();
                var translatedSingleEraseTarget = _statementTranslator.Translate(
                    singleEraseTarget.Target,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified,
                    _logger.Warning
                );
                var undeclaredVariablesInSingleEraseTarget = translatedSingleEraseTarget.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter).ToArray();
                foreach (var undeclaredVariable in undeclaredVariablesInSingleEraseTarget)
                    _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesInSingleEraseTarget);
                if (singleEraseTarget.ArgumentsIfAny == null)
                {
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "{0}.ERASE({1}, {2} => {{ {1} = {2}; }});",
                            _supportRefName.Name,
                            translatedSingleEraseTarget.TranslatedContent,
                            _tempNameGenerator(new CSharpName("v"), scopeAccessInformation).Name
                        ),
                        indentationDepth
                    ));
                }
                else
                {
                    // Note: "Erase a()" is a runtime error condition - either "a" is an array, in which case it will be a "Subscript out of range" or "a" is a variable that
                    // is not an array, in which case it will be a "Type mismatch" (we verified earlier that "a" is in fact a variable - and not a function, for example).
                    // We have no choice but to let the ERASE function work this out at runtime (which the non-by-ref-argument signature will do).
                    var translatedArguments = singleEraseTarget.ArgumentsIfAny
                        .Select(argument => _statementTranslator.Translate(
                            argument,
                            scopeAccessInformation,
                            ExpressionReturnTypeOptions.NotSpecified,
                            _logger.Warning
                        ))
                        .ToArray(); // Going to evaluate everything twice, might as well ToArray it
                    translationResult = translationResult.Add(new TranslatedStatement(
                        string.Format(
                            "{0}.ERASE({1}{2}{3});",
                            _supportRefName.Name,
                            translatedSingleEraseTarget.TranslatedContent,
                            translatedArguments.Any() ? ", " : "",
                            string.Join(", ", translatedArguments.Select(a => a.TranslatedContent))
                        ),
                        indentationDepth
                    ));
                    var undeclaredVariablesInArguments = translatedArguments.SelectMany(arg => arg.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)).ToArray();
                    foreach (var undeclaredVariable in undeclaredVariablesInArguments)
                        _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                    translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesInArguments);
                }
            }
            if (byRefArgumentsToRewrite.Any())
            {
                indentationDepth -= numberOfIndentationLevelsToWithDrawAfterByRefArgumentsProcessed;
                translationResult = byRefArgumentsToRewrite.CloseByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
            }
            return translationResult;
        }
    }
}
