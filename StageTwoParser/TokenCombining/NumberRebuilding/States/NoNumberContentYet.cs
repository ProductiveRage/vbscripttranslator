using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class NoNumberContentYet : IAmLookingForNumberContent
    {
        public static NoNumberContentYet Instance { get { return new NoNumberContentYet(); } }
        private NoNumberContentYet() { }

        public TokenProcessResult Process(IToken token, PartialNumberContent numberContent)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            if (numberContent.Tokens.Any())
                throw new Exception("The numberContent reference should be empty when using the NoNumberContentYet processor");

            return new TokenProcessResult(
                numberContent,
                new[] { token },
                Common.GetDefaultProcessor(token)
            );
        }
    }
}
