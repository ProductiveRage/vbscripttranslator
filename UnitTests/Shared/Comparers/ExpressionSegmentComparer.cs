using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class ExpressionSegmentComparer : IEqualityComparer<ExpressionSegment>
    {
        public bool Equals(ExpressionSegment x, ExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var tokenSetComparer = new TokenSetComparer();
            if (!tokenSetComparer.Equals(x.MemberAccessTokens, y.MemberAccessTokens))
                return false;

            var expressSetComparer = new ExpressionSetComparer();
            if (!expressSetComparer.Equals(x.Arguments, y.Arguments))
                return false;

            return true;
        }

        public int GetHashCode(ExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
