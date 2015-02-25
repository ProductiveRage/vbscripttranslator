using System;
using System.Linq;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class IfBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
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
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
            if (logger == null)
                throw new ArgumentNullException("logger");

            _statementTranslator = statementTranslator;
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
            var numberOfAdditionalBlocksInjectedForErrorTrapping = 0;
            foreach (var conditionalEntry in ifBlock.ConditionalClauses.Select((conditional, index) => new { Conditional = conditional, Index = index }))
            {
                // This is the content that will ultimately be forced into a boolean and used to construct an "if" conditional, it it C# code. If there is no error-trapping
                // to worry about then we just have to wrap this in an IF call (the "IF" method in the runtime support class) and be done with it. If error-trapping MAY be
                // involved, though, then it's more complicated - we might be able to just use the IF extension method that takes an error registration token as an argument
                // (this is the second least-complicated code path) but if we're within a function (or property) and the there any function or property calls within this
                // translated content that take by-ref arguments and any of those arguments are by-ref arguments of the containing function, then it will have to be re-
                // written since C# will not allow "ref" arguments to be manipulated in lambas (and that is how by-ref arguments are dealt with when calling nested
                // functions or properties). This third arrangement is the most complicated code path.
                var conditionalContent = _statementTranslator.Translate(
                    conditionalEntry.Conditional.Condition,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified,
                    _logger.Warning
                );

                // Check whether error-trapping may come into play at runtime - if so then we either need to deal with errors or (worst case) rewrite the condition expression
                // to avoid trying to pass any ref arguments of the containing function / property (where applicable) as a by-ref argument into another function or property.
                // If any such rewriting is done then the values must be written back as soon as the condition evaluation ends (even if it errors).
                if (scopeAccessInformation.MayRequireErrorWrapping(ifBlock))
                {
                    var byRefArgumentIdentifier = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
                    var byRefArgumentsToRewrite = byRefArgumentIdentifier.GetByRefArgumentsThatNeedRewriting(
                        conditionalEntry.Conditional.Condition.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
                        scopeAccessInformation,
                        new NonNullImmutableList<FuncByRefMapping>()
                    );
                    if (!byRefArgumentsToRewrite.Any())
                    {
                        // If we're not in a function or property or if that function or property does not have any by-ref arguments that we need to pass in as by-ref arguments
                        // to further functions or properties, then we're in the less complicate error-trapping scenario; we only have to use the IF extension method that deals
                        // with error-trapping.
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
                    else
                    {
                        // If we're in a function or property and that function / property has by-ref arguments that we then need to pass into further function / property calls
                        // in order to evaluate the current conditional, then we need to record those values in temporary references, try to evaulate the condition and then push
                        // the temporary values back into the original references. This is required in order to be consistent with VBScript and yet also produce code that compiles
                        // as C# (which will not let "ref" arguments of the containing function be used in lambdas, which is how we deal with updating by-ref arguments after function
                        // or property calls complete).
                        var evaluatedResultName = _tempNameGenerator(new CSharpName("ifResult"), scopeAccessInformation);
                        scopeAccessInformation = scopeAccessInformation.ExtendVariables(
                            byRefArgumentsToRewrite
                                .Select(r => new ScopedNameToken(r.To.Name, r.From.LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith))
                                .ToNonNullImmutableList()
                        );
                        translationResult = translationResult.Add(new TranslatedStatement(
                            "bool " + evaluatedResultName.Name + ";",
                            indentationDepth
                        ));
                        translationResult = byRefArgumentsToRewrite.OpenByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
                        indentationDepth++;
                        var rewrittenConditionalContent = _statementTranslator.Translate(
                            byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(conditionalEntry.Conditional.Condition, _nameRewriter),
                            scopeAccessInformation,
                            ExpressionReturnTypeOptions.NotSpecified,
                            _logger.Warning
                        );
                        translationResult = translationResult.Add(new TranslatedStatement(
                            string.Format(
                                "{0} = {1}.IF(() => {2}, {3});",
                                evaluatedResultName.Name,
                                _supportRefName.Name,
                                rewrittenConditionalContent.TranslatedContent,
                                scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                            ),
                            indentationDepth
                        ));
                        indentationDepth--;
                        translationResult = byRefArgumentsToRewrite.CloseByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
                        conditionalContent = new TranslatedStatementContentDetails(
                            evaluatedResultName.Name,
                            rewrittenConditionalContent.VariablesAccessed
                        );
                    }
                }
                else
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

                var undeclaredVariablesAccessed = conditionalContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
                foreach (var undeclaredVariable in undeclaredVariablesAccessed)
                    _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
                translationResult = translationResult.AddUndeclaredVariables(undeclaredVariablesAccessed);
                var innerStatements = conditionalEntry.Conditional.Statements.ToNonNullImmutableList();
                var conditionalInlineCommentIfAny = !innerStatements.Any() ? null : (innerStatements.First() as InlineCommentStatement);
                if (conditionalInlineCommentIfAny != null)
                    innerStatements = innerStatements.RemoveAt(0);
                translationResult = translationResult.Add(
                    new TranslatedStatement(
                        string.Format(
                            "{0} ({1}){2}",
                            (conditionalEntry.Index == 0) || scopeAccessInformation.MayRequireErrorWrapping(ifBlock) ? "if" : "else if",
                            conditionalContent.TranslatedContent,
                            (conditionalInlineCommentIfAny == null) ? "" : (" //" + conditionalInlineCommentIfAny.Content)
                        ),
                        indentationDepth
                    )
                );
                translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                translationResult = translationResult.Add(
                    Translate(
                        innerStatements,
                        scopeAccessInformation.SetParent(ifBlock),
                        indentationDepth + 1
                    )
                );
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));

                // If we're dealing with multiple if (a).. elseif (b).. elseif (c).. [else] blocks then these would be most simply represented by if (a).. else if (b)..
                // else if (c).. [else] blocks in C#. However, if error-trapping is involved then some of the conditions may have to be rewritten to deal with by-ref arguments
                // and then (after the condition is evaluated) those rewritten arguments need to be reflected on the original references (see the note further up for more
                // details). In this case, each subsequent "if" condition must be within its own "else" block in order for the the rewritten condition to be evaulated when
                // required (and not before). When this happens, there will be greater levels of nesting required. This nesting is injected here and tracked with the variable
                // "numberOfAdditionalBlocksInjectedForErrorTrapping" (this will be used further down to ensure that any extra levels of nesting are closed off).
                if (scopeAccessInformation.MayRequireErrorWrapping(ifBlock) && (conditionalEntry.Index < (ifBlock.ConditionalClauses.Count() - 1)))
                {
                    translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth));
                    translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                    indentationDepth++;
                    numberOfAdditionalBlocksInjectedForErrorTrapping++;
                }
            }

            if (ifBlock.OptionalElseClause != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement("else", indentationDepth));
                translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                translationResult = translationResult.Add(
                    Translate(
                        ifBlock.OptionalElseClause.Statements.ToNonNullImmutableList(),
                        scopeAccessInformation.SetParent(ifBlock),
                        indentationDepth + 1
                    )
                );
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }

            // If any additional levels of nesting were required above (for error-trapping scenarios), ensure they are closed off here
            for (var index = 0; index < numberOfAdditionalBlocksInjectedForErrorTrapping; index++)
            {
                indentationDepth--;
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
