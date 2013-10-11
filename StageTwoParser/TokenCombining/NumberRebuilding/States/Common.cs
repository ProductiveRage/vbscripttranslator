using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public static class Common
    {
        /// <summary>
        /// If a processor has encountered a token that has invalidated the assumptions it was making that it would be able to rebuild a number
        /// token, then it can pass the current token and PartialNumberContent to this to get back a TokenProcessResult which will return the
        /// unprocessed token content and select an appropriate processor to restart with.
        /// </summary>
        public static TokenProcessResult Reset(IToken token, PartialNumberContent numberContent)
        {
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");
            if (token == null)
                throw new ArgumentNullException("token");

            return new TokenProcessResult(
                new PartialNumberContent(new IToken[0]),
                numberContent.Tokens.Concat(new[] { token }),
                GetDefaultProcessor(token)
            );
        }

        public static IAmLookingForNumberContent GetDefaultProcessor(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return token.CouldPrecedeDecimalPointOrNegativeSign()
                ? (IAmLookingForNumberContent)PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber.Instance
                : NoNumberContentYet.Instance;
        }
    }
}
