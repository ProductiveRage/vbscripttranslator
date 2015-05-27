using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class ExpressionSegmentComparer : IEqualityComparer<IExpressionSegment>
    {
        public bool Equals(IExpressionSegment x, IExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            if (x.GetType() != y.GetType())
                return false;

            if (x.GetType() == typeof(BracketedExpressionSegment))
                return (new BracketedExpressionSegmentComparer()).Equals((BracketedExpressionSegment)x, (BracketedExpressionSegment)y);
            else if (x.GetType() == typeof(BuiltInValueExpressionSegment))
                return (new BuiltInValueExpressionSegmentComparer()).Equals((BuiltInValueExpressionSegment)x, (BuiltInValueExpressionSegment)y);
            else if (x.GetType() == typeof(OperationExpressionSegment))
                return (new OperatorOrComparisonExpressionSegmentComparer()).Equals((OperationExpressionSegment)x, (OperationExpressionSegment)y);
            else if (x.GetType() == typeof(NewInstanceExpressionSegment))
                return (new NewInstanceExpressionSegmentComparer()).Equals((NewInstanceExpressionSegment)x, (NewInstanceExpressionSegment)y);
            else if (x.GetType() == typeof(CallExpressionSegment))
                return (new CallSetItemExpressionSegmentComparer()).Equals((CallExpressionSegment)x, (CallExpressionSegment)y);
            else if (x.GetType() == typeof(CallSetExpressionSegment))
                return (new CallSetExpressionSegmentComparer()).Equals((CallSetExpressionSegment)x, (CallSetExpressionSegment)y);
            else if (x.GetType() == typeof(NumericValueExpressionSegment))
                return (new NumericValueExpressionSegmentComparer()).Equals((NumericValueExpressionSegment)x, (NumericValueExpressionSegment)y);
            else if (x.GetType() == typeof(DateValueExpressionSegment))
                return (new DateValueExpressionSegmentComparer()).Equals((DateValueExpressionSegment)x, (DateValueExpressionSegment)y);
            else if (x.GetType() == typeof(StringValueExpressionSegment))
                return (new StringValueExpressionSegmentComparer()).Equals((StringValueExpressionSegment)x, (StringValueExpressionSegment)y);
            else if (x.GetType() == typeof(RuntimeErrorExpressionSegment))
                return (new RuntimeErrorExpressionSegmentComparer()).Equals((RuntimeErrorExpressionSegment)x, (RuntimeErrorExpressionSegment)y);
            else
                throw new NotSupportedException("Unsupported IExpressionSegment type: " + x.GetType());
        }

        public int GetHashCode(IExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
