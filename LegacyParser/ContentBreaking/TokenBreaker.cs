using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.ContentBreaking
{
    public static class TokenBreaker
    {
        private static string WhiteSpaceChars = new string(
            Enumerable.Range((int)char.MinValue, (int)char.MaxValue).Select(v => (char)v).Where(c => char.IsWhiteSpace(c)).ToArray()
        );

        private const string TokenBreakChars = "_,.*&+-=!(){}[]:;\n";

        /// <summary>
        /// Break down an UnprocessedContentToken into a combination of AtomToken and AbstractEndOfStatementToken references. This will never
        /// return null nor a set containing any null references.
        /// </summary>
        public static IEnumerable<IToken> BreakUnprocessedToken(UnprocessedContentToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            var buffer = "";
            var content = token.Content;
            var tokens = new List<IToken>();
            for (var index = 0; index < content.Length; index++)
            {
                var chr = content.Substring(index, 1);
                if ((chr != "\n") && WhiteSpaceChars.IndexOf(chr) != -1)
                {
                    // If we've found a (non-line-return) whitespace character, push content
                    // retrieved from the token so far (if any), into a fresh token on the
                    // list and clear the buffer to accept following data.
                    if (buffer != "")
                        tokens.Add(AtomToken.GetNewToken(buffer));
                    buffer = "";
                }
                else if (TokenBreakChars.IndexOf(chr) != -1)
                {
                    // If we've found another "break" character (which means a token split
                    // is identified, but that we want to keep the break character itself,
                    // unlike with whitespace breaks), then do similar to above.
                    if (buffer != "")
                        tokens.Add(AtomToken.GetNewToken(buffer));
                    tokens.Add(AtomToken.GetNewToken(chr));
                    buffer = "";
                }
                else
                    buffer += chr;
            }
            if (buffer != "")
                tokens.Add(AtomToken.GetNewToken(buffer));
            
            // Handle ignore-line-return / end-of-statement combinations
            tokens = handleLineReturnCancels(tokens);

            return tokens;
        }

        /// <summary>
        /// Look for any "_" character AtomTokens and ensure they are followed by a line
        /// return - if so, drop both (if not, raise exception - invalid VBScript)
        /// </summary>
        private static List<IToken> handleLineReturnCancels(List<IToken> tokens)
        {
            var tokensOut = new List<IToken>();
            for (int index = 0; index < tokens.Count; index++)
            {
                var token = tokens[index];
                if ((token is AtomToken) && (token.Content == "_"))
                {
                    // Ensure followed by line return, then ignore both tokens
                    if (index == (tokens.Count - 1))
                        throw new Exception("Encountered line-return cancellation that isn't followed by a line return - invalid");
                    var tokenNext = tokens[index + 1];
                    if (!(tokenNext is EndOfStatementNewLineToken))
                        throw new Exception("Encountered line-return cancellation that isn't followed by a line return - invalid");
                    index++;
                }
                else
                    tokensOut.Add(token);
            }
            return tokensOut;
        }
    }
}
