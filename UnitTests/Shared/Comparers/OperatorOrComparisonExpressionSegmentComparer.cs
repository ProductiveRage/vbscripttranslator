using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class OperatorOrComparisonExpressionSegmentComparer : IEqualityComparer<OperatorOrComparisonExpressionSegment>
    {
        public bool Equals(OperatorOrComparisonExpressionSegment x, OperatorOrComparisonExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var tokenComparer = new TokenComparer();
            return tokenComparer.Equals(x.Token, y.Token);
        }

        public int GetHashCode(OperatorOrComparisonExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
