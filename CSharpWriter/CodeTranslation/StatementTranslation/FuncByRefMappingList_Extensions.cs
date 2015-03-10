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
        /// Where variables must be stored in tempoarar - TODO
        /// </summary>
        /// <param name="translationResult"></param>
        /// <param name="indentationDepth"></param>
        /// <param name="byRefArgumentsToRewrite"></param>
        /// <returns></returns>
        public static ByRefReplacementTranslationResultDetails OpenByRefReplacementDefinitionWork(
            this NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite,
            TranslationResult translationResult,
            int indentationDepth,
            VBScriptNameRewriter nameRewriter)
        {
            if (byRefArgumentsToRewrite == null)
                throw new ArgumentNullException("byRefArgumentsToRewrite");
            if (!byRefArgumentsToRewrite.Any())
                throw new ArgumentException("This should not be called without any byRefArgumentsToRewrite values since there is no pointing wrapping the work up in an additional try..finally in that case");
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            translationResult = translationResult.Add(new TranslatedStatement(
                string.Format(
                    "object {0};",
                    string.Join(
                        ", ",
                        byRefArgumentsToRewrite.Select(r => r.To.Name + " = " + nameRewriter(r.From).Name)
                    )
                ),
                indentationDepth
            ));

            if (byRefArgumentsToRewrite.All(mapping => mapping.MappedValueIsReadOnly))
                return new ByRefReplacementTranslationResultDetails(translationResult, distanceToIndentCodeWithMappedValues: 0);

            return new ByRefReplacementTranslationResultDetails(
                translationResult
                    .Add(new TranslatedStatement("try", indentationDepth))
                    .Add(new TranslatedStatement("{", indentationDepth)),
                distanceToIndentCodeWithMappedValues: 1
            );
        }

        public static Statement RewriteStatementUsingByRefArgumentMappings(this NonNullImmutableList<FuncByRefMapping> byRefArgumentsToRewrite, Statement statementBlock, VBScriptNameRewriter nameRewriter)
        {
            if (byRefArgumentsToRewrite == null)
                throw new ArgumentNullException("byRefArgumentsToRewrite");
            if (!byRefArgumentsToRewrite.Any())
                throw new ArgumentException("This should not be called without any byRefArgumentsToRewrite values since there is no pointing wrapping the work up in an additional try..finally in that case");
            if (statementBlock == null)
                throw new ArgumentNullException("statementBlock");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

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
            if (!byRefArgumentsToRewrite.Any())
                throw new ArgumentException("This should not be called without any byRefArgumentsToRewrite values since there is no pointing wrapping the work up in an additional try..finally in that case");
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

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
            if (!byRefArgumentsToRewrite.Any())
                throw new ArgumentException("This should not be called without any byRefArgumentsToRewrite values since there is no pointing wrapping the work up in an additional try..finally in that case");
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            if (byRefArgumentsToRewrite.All(mapping => mapping.MappedValueIsReadOnly))
                return translationResult;

            return translationResult
                .Add(new TranslatedStatement("}", indentationDepth))
                .Add(new TranslatedStatement(
                    string.Format(
                        "finally {{ {0}; }}",
                        string.Join(
                            "; ",
                            byRefArgumentsToRewrite.Select(r => nameRewriter(r.From).Name + " = " + r.To.Name)
                        )
                    ),
                    indentationDepth
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