using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class GotMinusSignOfNumber : IAmLookingForNumberContent
    {
        public static GotMinusSignOfNumber Instance { get { return new GotMinusSignOfNumber(); } }
        private GotMinusSignOfNumber() { }

        public TokenProcessResult Process(IToken token, PartialNumberContent numberContent)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            // At this point, the current token needs to be either a number or a decimal point. Otherwise it's not going to be
            // part of a valid numeric value.
            if (token is NumericValueToken)
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotSomeIntegerNumberContent.Instance
                );
            }
            else if (token.Is<MemberAccessorOrDecimalPointToken>())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotSomeDecimalNumberContent.Instance
                );
            }
            return Common.Reset(token, numberContent);
        }
    }
}
