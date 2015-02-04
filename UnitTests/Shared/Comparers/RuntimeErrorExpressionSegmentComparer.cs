using System;
using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class RuntimeErrorExpressionSegmentComparer : IEqualityComparer<RuntimeErrorExpressionSegment>
    {
        public bool Equals(RuntimeErrorExpressionSegment x, RuntimeErrorExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            return
                (x.RenderedContent == y.RenderedContent) &&
                (x.ExceptionType == y.ExceptionType) &&
                (x.Message == y.Message) &&
                new TokenSetComparer().Equals(x.AllTokens, y.AllTokens);
        }

        public int GetHashCode(RuntimeErrorExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
