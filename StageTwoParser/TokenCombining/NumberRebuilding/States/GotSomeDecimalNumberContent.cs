using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public class GotSomeDecimalNumberContent : IAmLookingForNumberContent
    {
        public static GotSomeDecimalNumberContent Instance { get { return new GotSomeDecimalNumberContent(); } }
        private GotSomeDecimalNumberContent() { }

        public TokenProcessResult Process(IEnumerable<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            // If we've already got tokens describing a decimal number and we encounter another decimal point then things are going badly
            // (this can't represent a valid number)
            if (token.Is<MemberAccessorOrDecimalPointToken>())
                throw new Exception("Encountered a MemberAccessorOrDecimalPointToken while part way through processing a decimal value - invalid content");

            // If we hit a numeric token, though, then things ARE going well and we can conclude the number search and incorporate the current token
            if (token is NumericValueToken)
            {
                var number = numberContent.AddToken(token).TryToExpressNumberFromTokens();
                if (number == null)
                    throw new Exception("numberContent should describe a number, null was returned from TryToExpressNumberFromTokens - invalid content");
                return new TokenProcessResult(
                    new PartialNumberContent(),
                    new[] { new NumericValueToken(number.Value, numberContent.Tokens.First().LineIndex) },
                    Common.GetDefaultProcessor(tokens)
                );
            }
            else
            {
                // If we hit any other token then hopefully things have gone well and we extracted a number, but not have processed the current
                // token (we don't have to try to process it as number content since it's not valid for two number tokens to exist adjacently)
                var number = numberContent.TryToExpressNumberFromTokens();
                if (number == null)
                {
                    if ((numberContent.Tokens.Count() == 1) && numberContent.Tokens.Single().Is<MemberAccessorOrDecimalPointToken>())
                    {
                        // If we've hit what appears to be a decimal point, but it's the current token does not result in a successful numeric
                        // parse, then presumably it's actually a property or method access within a "WITH" construct. As such, don't try to
                        // treat it as numeric content - return the tokens as "processed" (meaning there is no NumberRebuilder processing
                        // to be done to them).
                        return new TokenProcessResult(
                            new PartialNumberContent(),
                            new[] { numberContent.Tokens.Single(), token },
                            Common.GetDefaultProcessor(tokens)
                        );
                    }
                    else
                        throw new Exception("numberContent should describe a number, null was returned from TryToExpressNumberFromTokens - invalid content");
                }
                return new TokenProcessResult(
                    new PartialNumberContent(),
                    new[] { new NumericValueToken(number.Value, numberContent.Tokens.First().LineIndex), token },
                    Common.GetDefaultProcessor(tokens)
                );
            }
        }
    }
}
