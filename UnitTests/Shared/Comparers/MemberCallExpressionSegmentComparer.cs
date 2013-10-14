using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class MemberCallExpressionSegmentComparer : IEqualityComparer<CallExpressionSegment>
    {
        public bool Equals(CallExpressionSegment x, CallExpressionSegment y)
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

        public int GetHashCode(CallExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
