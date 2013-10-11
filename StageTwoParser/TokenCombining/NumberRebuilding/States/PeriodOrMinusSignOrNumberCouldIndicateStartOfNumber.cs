using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber : IAmLookingForNumberContent
    {
        public static PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber Instance { get { return new PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber(); } }
        private PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber() { }

        public TokenProcessResult Process(IToken token, PartialNumberContent numberContent)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

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
            else if (token.IsNumericAtomToken())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotSomeIntegerNumberContent.Instance
                );
            }

            return Common.Reset(token, numberContent);
        }
    }
}
