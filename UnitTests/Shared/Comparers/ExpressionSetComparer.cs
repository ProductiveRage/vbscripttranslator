using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class ExpressionSetComparer : IEqualityComparer<IEnumerable<Expression>>
    {
        public bool Equals(IEnumerable<Expression> x, IEnumerable<Expression> y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var arrayX = x.ToArray();
            var arrayY = y.ToArray();
            if (arrayX.Length != arrayY.Length)
                return false;

            var expressionComparer = new ExpressionComparer();
            for (var index = 0; index < arrayX.Length; index++)
            {
                if (!expressionComparer.Equals(arrayX[index], arrayY[index]))
                    return false;
            }
            return true;
        }

        public int GetHashCode(IEnumerable<Expression> obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
