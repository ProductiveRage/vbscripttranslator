using System;
using System.Linq;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
    public static class FuncByRefMappingList_Extensions
    {
        /// <summary>
        /// Where variables must be stored in temporary references in order to be accessed within lambas (which is the case for variables that are "ref" arguments of the containing function, the common cases
        /// where lambdas may be required are for passing as REF into an IProvideCallArguments implementation or when accessed within a HANDLEERROR call), the temporary references must be defined and a try
        /// opened before the work attempted. After the work is completed, in a finally, the aliases values must be mapped back onto the source values - this is what the CloseByRefReplacementDefinitionWork
        /// method is for.
        /// </summary>
        public static ByRefReplacementTranslationResultDetails OpenByRefReplacementDefinitionWork(
            this NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite,
            TranslationResult translationResult,
            int indentationDepth,
            VBScriptNameRewriter nameRewriter)
        {
            if (byRefArgumentsToRewrite == null)
                throw new ArgumentNullException("byRefArgumentsToRewrite");
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            // Originally, this would throw an exception if there were no by-ref arguments (why bother calling this if there are no by-ref arguments to deal with; does this indicate an error in the calling
            // code?) but in some cases it's easier to be able to call it without having check whether there were any value that need rewriting and the cases where being so strict may catch unintentional
            // calls are few
            if (!byRefArgumentsToRewrite.Any())
                return new ByRefReplacementTranslationResultDetails(translationResult, 0);

			var lineIndexForStartOfContent = byRefArgumentsToRewrite.Min(a => a.From.LineIndex);
			translationResult = translationResult.Add(new TranslatedStatement(
                string.Format(
                    "object {0};",
                    string.Join(
                        ", ",
                        byRefArgumentsToRewrite.Select(r => r.To.Name + " = " + nameRewriter(r.From).Name)
                    )
                ),
                indentationDepth,
				lineIndexForStartOfContent
			));

            if (byRefArgumentsToRewrite.All(mapping => mapping.MappedValueIsReadOnly))
                return new ByRefReplacementTranslationResultDetails(translationResult, distanceToIndentCodeWithMappedValues: 0);

            return new ByRefReplacementTranslationResultDetails(
                translationResult
                    .Add(new TranslatedStatement("try", indentationDepth, lineIndexForStartOfContent))
                    .Add(new TranslatedStatement("{", indentationDepth, lineIndexForStartOfContent)),
                distanceToIndentCodeWithMappedValues: 1
            );
        }

        public static Statement RewriteStatementUsingByRefArgumentMappings(this NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite, Statement statementBlock, VBScriptNameRewriter nameRewriter)
        {
            if (byRefArgumentsToRewrite == null)
                throw new ArgumentNullException("byRefArgumentsToRewrite");
            if (statementBlock == null)
                throw new ArgumentNullException("statementBlock");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            // Originally, this would throw an exception if there were no by-ref arguments (why bother calling this if there are no by-ref arguments to deal with; does this indicate an error in the calling
            // code?) but in some cases it's easier to be able to call it without having check whether there were any value that need rewriting and the cases where being so strict may catch unintentional
            // calls are few
            if (!byRefArgumentsToRewrite.Any())
                return statementBlock;

            return new Statement(
                statementBlock.Tokens.Select(t =>
                {
                    var nameToken = t as NameToken;
                    if (nameToken == null)
                        return t;
                    var referenceRewriteDetailsIfApplicable = byRefArgumentsToRewrite.FirstOrDefault(
                        r => nameRewriter.GetMemberAccessTokenName(r.From) == nameRewriter.GetMemberAccessTokenName(nameToken)
                    );
                    return (referenceRewriteDetailsIfApplicable == null) ? t : new DoNotRenameNameToken(referenceRewriteDetailsIfApplicable.To.Name, t.LineIndex);
                }),
                statementBlock.CallPrefix
            );
        }

        public static Expression RewriteExpressionUsingByRefArgumentMappings(this NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite, Expression expression, VBScriptNameRewriter nameRewriter)
        {
            if (byRefArgumentsToRewrite == null)
                throw new ArgumentNullException("byRefArgumentsToRewrite");
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            // Originally, this would throw an exception if there were no by-ref arguments (why bother calling this if there are no by-ref arguments to deal with; does this indicate an error in the calling
            // code?) but in some cases it's easier to be able to call it without having check whether there were any value that need rewriting and the cases where being so strict may catch unintentional
            // calls are few
            if (!byRefArgumentsToRewrite.Any())
                return expression;

            return new Expression(
                expression.Tokens.Select(t =>
                {
                    var nameToken = t as NameToken;
                    if (nameToken == null)
                        return t;
                    var referenceRewriteDetailsIfApplicable = byRefArgumentsToRewrite.FirstOrDefault(
                        r => nameRewriter.GetMemberAccessTokenName(r.From) == nameRewriter.GetMemberAccessTokenName(nameToken)
                    );
                    return (referenceRewriteDetailsIfApplicable == null) ? t : new DoNotRenameNameToken(referenceRewriteDetailsIfApplicable.To.Name, t.LineIndex);
                })
            );
        }

        public static TranslationResult CloseByRefReplacementDefinitionWork(
            this NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite,
            TranslationResult translationResult,
            int indentationDepth,
            VBScriptNameRewriter nameRewriter)
        {
            if (byRefArgumentsToRewrite == null)
                throw new ArgumentNullException("byRefArgumentsToRewrite");
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            // Originally, this would throw an exception if there were no by-ref arguments (why bother calling this if there are no by-ref arguments to deal with; does this indicate an error in the calling
            // code?) but in some cases it's easier to be able to call it without having check whether there were any value that need rewriting and the cases where being so strict may catch unintentional
            // calls are few
            if (!byRefArgumentsToRewrite.Any())
                return translationResult;

            if (byRefArgumentsToRewrite.All(mapping => mapping.MappedValueIsReadOnly))
                return translationResult;

			var lineIndexForEndOfContent = byRefArgumentsToRewrite.Max(a => a.From.LineIndex);
			return translationResult
                .Add(new TranslatedStatement("}", indentationDepth, lineIndexForEndOfContent))
                .Add(new TranslatedStatement(
                    string.Format(
                        "finally {{ {0}; }}",
                        string.Join(
                            "; ",
                            byRefArgumentsToRewrite.Select(r => nameRewriter(r.From).Name + " = " + r.To.Name)
                        )
					),
                    indentationDepth,
					lineIndexForEndOfContent
				));
        }

        public class ByRefReplacementTranslationResultDetails
        {
            public ByRefReplacementTranslationResultDetails(TranslationResult translationResult, int distanceToIndentCodeWithMappedValues)
            {
                if (translationResult == null)
                    throw new ArgumentNullException("translationResult");
                if (distanceToIndentCodeWithMappedValues < 0)
                    throw new ArgumentOutOfRangeException("distanceToIndentCodeWithMappedValues", "must be zero or greater");

                TranslationResult = translationResult;
                DistanceToIndentCodeWithMappedValues = distanceToIndentCodeWithMappedValues;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public TranslationResult TranslationResult { get; private set; }

            /// <summary>
            /// This will always be zero or greater
            /// </summary>
            public int DistanceToIndentCodeWithMappedValues { get; private set; }
        }
    }
}