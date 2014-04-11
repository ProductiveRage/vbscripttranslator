using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber : IAmLookingForNumberContent
    {
        public static PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber Instance { get { return new PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber(); } }
        private PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber() { }

        public TokenProcessResult Process(IEnumerable<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            if (token.Is<MemberAccessorOrDecimalPointToken>())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotSomeDecimalNumberContent.Instance
                );
            }
            else if (token.IsMinusSignOperator())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotMinusSignOfNumber.Instance
                );
            }
            else if (token is NumericValueToken)
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotSomeIntegerNumberContent.Instance
                );
            }

            return Common.Reset(tokens, numberContent);
        }
    }
}
