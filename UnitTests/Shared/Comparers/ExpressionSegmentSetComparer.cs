using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class ExpressionSegmentSetComparer : IEqualityComparer<IEnumerable<IExpressionSegment>>
    {
		public bool Equals(IEnumerable<IExpressionSegment> x, IEnumerable<IExpressionSegment> y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var segmentsX = x.ToArray();
            var segmentsY = y.ToArray();
            if (segmentsX.Length != segmentsY.Length)
                return false;

            var expressionSegmentComparer = new ExpressionSegmentComparer();
            for (var index = 0; index < segmentsX.Length; index++)
            {
                if (!expressionSegmentComparer.Equals(segmentsX[index], segmentsY[index]))
                    return false;
            }
            return true;
        }

		public int GetHashCode(IEnumerable<IExpressionSegment> obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
