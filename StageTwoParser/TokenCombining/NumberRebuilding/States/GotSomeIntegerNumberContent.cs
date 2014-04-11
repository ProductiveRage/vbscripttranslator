using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class GotSomeIntegerNumberContent : IAmLookingForNumberContent
    {
        public static GotSomeIntegerNumberContent Instance { get { return new GotSomeIntegerNumberContent(); } }
        private GotSomeIntegerNumberContent() { }

        public TokenProcessResult Process(IEnumerable<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            // The only continuation possibility for the number is if a decimal point is reached
            if (token.Is<MemberAccessorOrDecimalPointToken>())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    new IToken[0],
                    GotSomeDecimalNumberContent.Instance
                );
            }
            
            // If we're not at a decimal point then the end of the number content must have been reached
            // - Try to extract the number content so far and express that as a new token
            // - Return a "processedTokens" set of this and the current token (we don't need to worry about trying to process
            //   that here since it's not valid for two number tokens to exist adjacently with nothing in between)
            var number = numberContent.TryToExpressNumberFromTokens();
            if (number == null)
                throw new Exception("numberContent should describe a number, null was returned from TryToExpressNumberFromTokens - invalid content");
            return new TokenProcessResult(
                new PartialNumberContent(),
                new[] { new NumericValueToken(number.Value, numberContent.Tokens.First().LineIndex), token },
                Common.GetDefaultProcessor(tokens)
            );
        }
    }
}
