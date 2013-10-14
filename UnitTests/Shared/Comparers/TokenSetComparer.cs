using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class TokenSetComparer : IEqualityComparer<IEnumerable<IToken>>
    {
        public bool Equals(IEnumerable<IToken> x, IEnumerable<IToken> y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var tokensArrayX = x.ToArray();
            var tokensArrayY = y.ToArray();
            if (tokensArrayX.Length != tokensArrayY.Length)
                return false;

            var tokenComparer = new TokenComparer();
            for (var index = 0; index < tokensArrayX.Length; index++)
            {
                if (!tokenComparer.Equals(tokensArrayX[index], tokensArrayY[index]))
                    return false;
            }
            return true;
        }

        public int GetHashCode(IEnumerable<IToken> obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
