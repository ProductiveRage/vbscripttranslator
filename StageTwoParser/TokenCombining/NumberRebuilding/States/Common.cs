using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public static class Common
    {
        /// <summary>
        /// If a processor has encountered a token that has invalidated the assumptions it was making that it would be able to rebuild a number
        /// token, then it can pass the current tokens and PartialNumberContent to this to get back a TokenProcessResult which will return the
        /// unprocessed token content and select an appropriate processor to restart with.
        /// </summary>
        public static TokenProcessResult Reset(IEnumerable<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            return new TokenProcessResult(
                new PartialNumberContent(new IToken[0]),
                numberContent.Tokens.Concat(new[] { token }),
                GetDefaultProcessor(tokens)
            );
        }

        public static IAmLookingForNumberContent GetDefaultProcessor(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            return CouldPrecedeDecimalPointOrNegativeSign(token)
                ? (IAmLookingForNumberContent)PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber.Instance
                : NoNumberContentYet.Instance;
        }

        private static bool CouldPrecedeDecimalPointOrNegativeSign(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            // Note: The "Is" method requires a precise match, which is why Is<ComparisonOperatorToken> is required as well as Is<OperatorToken>
            return (token.Is<OpenBrace>() || token.Is<OperatorToken>() || token.Is<ComparisonOperatorToken>() || token.Is<KeyWordToken>());
        }
    }
}
