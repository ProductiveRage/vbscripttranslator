using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class DateValueExpressionSegmentComparer : IEqualityComparer<DateValueExpressionSegment>
    {
        public bool Equals(DateValueExpressionSegment x, DateValueExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            return x.Token.Content.Equals(y.Token.Content, StringComparison.InvariantCultureIgnoreCase);
        }

        public int GetHashCode(DateValueExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
