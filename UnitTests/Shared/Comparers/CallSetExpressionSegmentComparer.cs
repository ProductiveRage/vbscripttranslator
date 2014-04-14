using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class CallSetExpressionSegmentComparer : IEqualityComparer<CallSetExpressionSegment>
    {
        public bool Equals(CallSetExpressionSegment x, CallSetExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var callExpressionSegmentsX = x.CallExpressionSegments.ToArray();
            var callExpressionSegmentsY = y.CallExpressionSegments.ToArray();
            if (callExpressionSegmentsX.Length != callExpressionSegmentsY.Length)
                return false;

            var callExpressionComparer = new CallSetItemExpressionSegmentComparer();
            for (var index = 0; index < callExpressionSegmentsX.Length; index++)
            {
                if (!callExpressionComparer.Equals(callExpressionSegmentsX[index], callExpressionSegmentsY[index]))
                    return false;
            }
            return true;
        }

        public int GetHashCode(CallSetExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
