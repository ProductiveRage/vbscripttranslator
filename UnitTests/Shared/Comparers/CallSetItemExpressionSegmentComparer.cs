using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class CallSetItemExpressionSegmentComparer : IEqualityComparer<CallSetItemExpressionSegment>
    {
        public bool Equals(CallSetItemExpressionSegment x, CallSetItemExpressionSegment y)
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

			if (x.ZeroArgumentBracketsPresence != y.ZeroArgumentBracketsPresence)
				return false;

            return true;
        }

        public int GetHashCode(CallSetItemExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
