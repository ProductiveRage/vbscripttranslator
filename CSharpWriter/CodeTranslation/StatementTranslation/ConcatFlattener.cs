using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
    /// <summary>
    /// Runs of multiple string concatenations are a common occurrence in VBScript, so rather then making the translated code longer than it needs to be, by limiting
    /// the number of arguments takens by the CONCAT method to two (like the other operators - except for NOT, which takes only one argument), the CONCAT method may
    /// may also take more than two arguments if it would make no difference to the enforcing of operator precedence. This methods will change an expression that has
    /// been strictly parsed to allow only two CONCAT arguments and rearrange it to support more arguments for a single CONCAT call if it would have no other effect
    /// on the expression. If this manipulation is not possible then no change to the data will be performed.
    /// </summary>
    public static class ConcatFlattener
    {
        public static Expression Flatten(Expression expression)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");

            return new Expression(
                Flatten(expression.Segments)
            );
        }

        private static IEnumerable<IExpressionSegment> Flatten(IEnumerable<IExpressionSegment> expressionSegments)
        {
            if (expressionSegments == null)
                throw new ArgumentNullException("expressionSegments");

            if (!IsTwoValueConcat(expressionSegments))
                return expressionSegments;

            var expressionSegmentsArray = expressionSegments.ToArray();
            if (expressionSegmentsArray.Any(s => s == null))
                throw new ArgumentException("Null reference encountered in expressionSegments set");

            var flattenedSegments = new List<IExpressionSegment>();
            var firstSegmentAsBracketedExpressionSegment = expressionSegmentsArray[0] as BracketedExpressionSegment;
            if ((firstSegmentAsBracketedExpressionSegment != null) && IsTwoValueConcat(firstSegmentAsBracketedExpressionSegment.Segments))
                flattenedSegments.AddRange(Flatten(new Expression(firstSegmentAsBracketedExpressionSegment.Segments)).Segments);
            else
                flattenedSegments.Add(expressionSegmentsArray[0]);
            flattenedSegments.Add(expressionSegmentsArray[1]);
            var lastSegmentAsBracketedExpressionSegment = expressionSegmentsArray[2] as BracketedExpressionSegment;
            if ((lastSegmentAsBracketedExpressionSegment != null) && IsTwoValueConcat(lastSegmentAsBracketedExpressionSegment.Segments))
                flattenedSegments.AddRange(Flatten(new Expression(lastSegmentAsBracketedExpressionSegment.Segments)).Segments);
            else
                flattenedSegments.Add(expressionSegmentsArray[2] );
            return flattenedSegments;
        }

        private static bool IsTwoValueConcat(IEnumerable<IExpressionSegment> expressionSegments)
        {
            if (expressionSegments == null)
                throw new ArgumentNullException("expressionSegments");

            var expressionSegmentsArray = expressionSegments.ToArray();
            if (expressionSegmentsArray.Any(s => s == null))
                throw new ArgumentException("Null reference encountered in expressionSegments set");

            return (expressionSegmentsArray.Length == 3)
                && !(expressionSegmentsArray[0] is OperationExpressionSegment)
                && (expressionSegmentsArray[1] is OperationExpressionSegment)
                && (((OperationExpressionSegment)expressionSegmentsArray[1]).Token.Content == "&")
                && !(expressionSegmentsArray[2] is OperationExpressionSegment);
        }
    }
}
