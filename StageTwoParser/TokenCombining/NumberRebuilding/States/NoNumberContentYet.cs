using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class NoNumberContentYet : IAmLookingForNumberContent
    {
        public static NoNumberContentYet Instance { get { return new NoNumberContentYet(); } }
        private NoNumberContentYet() { }

        public TokenProcessResult Process(IEnumerable<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            if (numberContent.Tokens.Any())
                throw new Exception("The numberContent reference should be empty when using the NoNumberContentYet processor");

            if (token.Is<NumericValueToken>())
            {
                return new TokenProcessResult(
                    new PartialNumberContent(new[] { token }),
                    new IToken[0],
                    GotSomeIntegerNumberContent.Instance
                );
            }

            return new TokenProcessResult(
                numberContent,
                new[] { token },
                Common.GetDefaultProcessor(tokens)
            );
        }
    }
}
