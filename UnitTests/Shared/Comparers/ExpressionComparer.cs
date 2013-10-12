using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.StageTwoParser;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class ExpressionComparer : IEqualityComparer<Expression>
    {
        public bool Equals(Expression x, Expression y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var segmentsX = x.Segments.ToArray();
            var segmentsY = y.Segments.ToArray();
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

        public int GetHashCode(Expression obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
