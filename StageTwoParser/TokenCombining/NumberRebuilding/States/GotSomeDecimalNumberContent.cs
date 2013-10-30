using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class GotSomeDecimalNumberContent : IAmLookingForNumberContent
    {
        public static GotSomeDecimalNumberContent Instance { get { return new GotSomeDecimalNumberContent(); } }
        private GotSomeDecimalNumberContent() { }

        public TokenProcessResult Process(IToken token, PartialNumberContent numberContent)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            // If we've already got tokens describing a decimal number and we encounter another decimal point then things are going badly
            // (this can't represent a valid number)
            if (token.Is<MemberAccessorOrDecimalPointToken>())
                throw new Exception("Encountered a MemberAccessorOrDecimalPointToken while part way through processing a decimal value - invalid  content");

            // If we hit a numeric token, though, then things ARE going well and we can conclude the number search and incorporate the current token
            if (token is NumericValueToken)
            {
                var number = numberContent.AddToken(token).TryToExpressNumberFromTokens();
                if (number == null)
                    throw new Exception("numberContent should describe a number, null was returned from TryToExpressNumberFromTokens - invalid content");
                return new TokenProcessResult(
                    new PartialNumberContent(),
                    new[] { new NumericValueToken(number.Value, numberContent.Tokens.First().LineIndex) },
                    Common.GetDefaultProcessor(token)
                );
            }
            else
            {
                // If we hit any other token then hopefully things have gone well and we extract a number, but not have processed the current
                // token (we don't have to try to process it as number content since it's not valid for two number tokens to exist adjacently)
                var number = numberContent.TryToExpressNumberFromTokens();
                if (number == null)
                    throw new Exception("numberContent should describe a number, null was returned from TryToExpressNumberFromTokens - invalid content");
                return new TokenProcessResult(
                    new PartialNumberContent(),
                    new[] { new NumericValueToken(number.Value, numberContent.Tokens.First().LineIndex), token },
                    Common.GetDefaultProcessor(token)
                );
            }
        }
    }
}
