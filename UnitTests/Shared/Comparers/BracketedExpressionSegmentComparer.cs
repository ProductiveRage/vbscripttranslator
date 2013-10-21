using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class BracketedExpressionSegmentComparer : IEqualityComparer<BracketedExpressionSegment>
    {
        public bool Equals(BracketedExpressionSegment x, BracketedExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var expressionSegmentSetComparer = new ExpressionSegmentSetComparer();
			return expressionSegmentSetComparer.Equals(x.Segments, y.Segments);
        }

        public int GetHashCode(BracketedExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
