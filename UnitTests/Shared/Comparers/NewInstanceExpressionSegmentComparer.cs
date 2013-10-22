using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class NewInstanceExpressionSegmentComparer : IEqualityComparer<NewInstanceExpressionSegment>
    {
        public bool Equals(NewInstanceExpressionSegment x, NewInstanceExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            return x.ClassName.Content.Equals(y.ClassName.Content, StringComparison.InvariantCultureIgnoreCase);
        }

        public int GetHashCode(NewInstanceExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
