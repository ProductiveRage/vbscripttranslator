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

            // If this token is a NumericValueToken then this is the start of numeric content (since we don't currently have any number
            // content at all). At this point we only know that it's the start of an integer value, the GotSomeIntegerNumberContent
            // state will deal with any decimal points that are encountered (switching to the GotSomeDecimalNumberContent state).
            if (token.Is<NumericValueToken>())
            {
                return new TokenProcessResult(
                    new PartialNumberContent(new[] { token }),
                    new IToken[0],
                    GotSomeIntegerNumberContent.Instance
                );
            }

            // If this is a MemberAccessorOrDecimalPointToken then it could be a member accessor (eg. "a.Name") or it could be the
            // start of a zero-less decimal number (eg. "fnc .1") - if the next token is a NumericValueToken then it is the latter
            // case and we need to switch to the GotSomeDecimalNumberContent state.
            if (token.Is<MemberAccessorOrDecimalPointToken>())
            {
                var nextTokens = tokens.Skip(1);
                if (nextTokens.Any())
                {
                    var nextToken = nextTokens.First();
                    if (nextToken == null)
                        throw new ArgumentException("Null reference encountered in tokens set");
                    if (nextToken.Is<NumericValueToken>())
                    {
                        return new TokenProcessResult(
                            new PartialNumberContent(new[] { token }),
                            new IToken[0],
                            GotSomeDecimalNumberContent.Instance
                        );
                    }
                }
            }

            return new TokenProcessResult(
                numberContent,
                new[] { token },
                Common.GetDefaultProcessor(tokens)
            );
        }
    }
}
