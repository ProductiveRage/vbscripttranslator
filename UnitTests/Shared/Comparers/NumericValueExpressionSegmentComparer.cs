using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class NumericValueExpressionSegmentComparer : IEqualityComparer<NumericValueExpressionSegment>
    {
        public bool Equals(NumericValueExpressionSegment x, NumericValueExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            return x.Token.Content.Equals(y.Token.Content, StringComparison.InvariantCultureIgnoreCase);
        }

        public int GetHashCode(NumericValueExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
