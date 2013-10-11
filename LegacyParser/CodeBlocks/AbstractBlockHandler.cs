using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks
{
    public abstract class AbstractBlockHandler
    {
        // =======================================================================================
        // ABSTRACT METHODS
        // =======================================================================================
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public abstract ICodeBlock Process(List<IToken> tokens);

        // =======================================================================================
        // HELPER METHODS FOR DERIVED CLASSES
        // =======================================================================================
        /// <summary>
        /// Grab specific token from list. Optionally specify that it must be an AtomToken in
        /// order to be valid. Will raise an exception if there are no more tokens available,
        /// or if a AtomToken was required but the next token was of a different type.
        /// </summary>
        protected IToken getToken(IEnumerable<IToken> tokens, int offset, IEnumerable<Type> allowedTokenTypes)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (offset < 0)
                throw new ArgumentException("Negative offset specified - invalid");
            if (offset >= tokens.Count())
                throw new ArgumentException("Insufficient tokens - invalid");
            if ((allowedTokenTypes != null) && !allowedTokenTypes.Any())
                    throw new ArgumentException("No allowed tokens types (pass as null to set no restriction");
            var token = tokens.ElementAt(offset);
            if (allowedTokenTypes != null)
            {
                bool validTokenType = false;
                foreach (var allowedType in allowedTokenTypes)
                {
                    if (isObjectOfTypeOrDerivedFrom(token, allowedType))
                    {
                        validTokenType = true;
                        break;
                    }
                }
                if (!validTokenType)
                    throw new Exception("Token is not of an allowed type [" + token.GetType().ToString()+ "]");
            }
            return token;
        }

        private bool isObjectOfTypeOrDerivedFrom(object obj, Type type)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");
            if (type == null)
                throw new ArgumentNullException("type");
            var objType = obj.GetType();
            while (true)
            {
                if (objType == type)
                    return true;
                if (objType.BaseType == null)
                    return false;
                objType = objType.BaseType;
            }
        }

        protected IToken getToken_AtomOnly(IEnumerable<IToken> tokens, int offset)
        {
            return getToken(tokens, offset, new List<Type>()
            {
                typeof(AtomToken)
            });
        }

        protected IToken getToken_AtomOrStringOnly(IEnumerable<IToken> tokens, int offset)
        {
            return getToken(tokens, offset, new List<Type>()
            {
                typeof(AtomToken),
                typeof(StringToken)
            });
        }

        protected bool isEndOfStatement(IEnumerable<IToken> tokens, int offset)
        {
            var token = getToken(tokens, offset, null);
            return (token is AbstractEndOfStatementToken);
        }

        /// <summary>
        /// Try to match AtomToken pattern - if there are insufficient tokens to match, or if a
        /// non-AtomToken is encountered, return false. Only rase exceptions if null tokens are
        /// found in the stream, the stream is null, the "values" array is null or empty of the
        /// optional offset value is less than zero. (If the offset value is too far along for
        /// the content to be matched, false will be returned).
        /// </summary>
        protected bool checkAtomTokenPattern(IEnumerable<IToken> tokens, string[] values, bool matchCase)
        {
            if (tokens == null)
                throw new ArgumentNullException("token");
            if (values == null)
                throw new ArgumentNullException("values");
            if (values.Length == 0)
                throw new ArgumentException("Zero values to match");
            foreach (string value in values)
            {
                if (value == null)
                    throw new ArgumentException("Null value specified");
            }
            var tokenArray = tokens.ToArray();
            if (tokenArray.Length < values.Length)
                return false;
            for (int index = 0; index < values.Length; index++)
            {
                // Only consider AtomTokens (if get anything else, we can't handle it)
                IToken token = tokenArray[index];
                if (token == null)
                    throw new ArgumentException("Null token specified");
                if (!(token is AtomToken))
                    return false;
                
                string val1 = values[index];
                string val2 = token.Content;
                if (!matchCase)
                {
                    val1 = val1.ToUpper();
                    val2 = val2.ToUpper();
                }
                if (val1 != val2)
                    return false;
            }
            return true;
        }

        protected bool checkAtomTokenPattern(IEnumerable<IToken> tokens, int offset, string[] values, bool matchCase)
        {
            // Validate input - throw exception if conditions not met
            if (tokens == null)
                throw new ArgumentNullException("token");
            if (values == null)
                throw new ArgumentNullException("values");
            if (values.Length == 0)
                throw new ArgumentException("Zero values to match");
            if (offset < 0)
                throw new ArgumentException("Invalid offset value < 0 [" + offset.ToString() + "]");

            // Try to grab section of token stream - if there are insufficient tokens, return
            // false rather than throwing an exception (this method is supposed to be flexible)
            var tokensArray = tokens.ToArray();
            if (offset + values.Length > tokensArray.Length)
                return false;
            return checkAtomTokenPattern(
                getTokenListSection(tokensArray, offset, values.Length),
                values,
                matchCase
            );
        }

        /// <summary>
        /// Extract a comma-separated sequence of values from a token stream, starting at the
        /// specified location. Continue until find a token matching the endMarker (both the
        /// token type and content must match). The endMarker will only be checked for after
        /// validated content - eg. if FunctionHandler needs to traverse parameter tokens with
        /// a ")" AtomToken endMarker, any "(", ")" sequences that complement each other will
        /// not count towards the endMarker. Only AtomTokens and StringTokens are permissible 
        /// in the token stream that are to be handled here, with the exception of the optional
        /// use of an EndOfStatementToken for the endMarker.
        /// </summary>
        protected List<List<IToken>> getEntryList(IEnumerable<IToken> tokens, int offset, IToken endMarker)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (offset < 0)
                throw new Exception("Negative offset specified - invalid");
            if (offset >= tokens.Count())
                throw new Exception("Insufficient tokens - invalid");
            if (endMarker == null)
                throw new ArgumentNullException("endMarker");
            if ((!(endMarker is AtomToken)) && (!(endMarker is AbstractEndOfStatementToken)))
                throw new ArgumentException("Invalid endMarker - must be Atom or EndOfStatement Token");

            var allowedTokenTypes = new List<Type>() { typeof(AtomToken), typeof(StringToken) };
            var entryList = new List<List<IToken>>();
            var buffer = new List<IToken>();
            var bracketCount = 0;
            while (true)
            {
                // Only check for endMarker if not in bracket sequence
                if (bracketCount == 0)
                {
                    // Check for endMarker
                    bool reachedEndMarker = false;
                    if ((endMarker is AbstractEndOfStatementToken) && isEndOfStatement(tokens, offset))
                        reachedEndMarker = true;
                    else
                    {
                        var possibleEndMarker = getToken(tokens, offset, allowedTokenTypes);
                        reachedEndMarker =
                            ((possibleEndMarker is AtomToken)
                            && (possibleEndMarker.Content.ToUpper() == endMarker.Content.ToUpper()));
                    }
                    if (reachedEndMarker)
                        break;
                }

                // Only check for separator if not in bracket sequence
                var gotSeparator = false;
                if (bracketCount == 0)
                {
                    IToken token = getToken(tokens, offset, allowedTokenTypes);
                    if (token.Content == ",")
                    {
                        // Got it.. add current entry to list (don't worry if it's blank,
                        // let the caller decide whether that's valid or not)
                        gotSeparator = true;
                        entryList.Add(buffer);
                        buffer = new List<IToken>();
                    }
                }
                if (!gotSeparator)
                {
                    // Not got separator, add to buffer (check for brackets)
                    IToken token = getToken(tokens, offset, allowedTokenTypes);
                    buffer.Add(token);
                    if (token.Content == "(")
                        bracketCount++;
                    else if (token.Content == ")")
                    {
                        bracketCount--;
                        if (bracketCount < 0)
                            throw new Exception("Invalid bracketing sequence");
                    }
                }
                offset++;
            }
            if (buffer.Count != 0)
                entryList.Add(buffer);
            return entryList;
        }

        /// <summary>
        /// Return a new list that is a subset of the input token list
        /// </summary>
        protected IEnumerable<IToken> getTokenListSection(IEnumerable<IToken> tokens, int start, int count)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            var tokensArray = tokens.ToArray();
            if ((start < 0) || (start >= tokensArray.Length))
                throw new ArgumentException("Invalid start value [" + start.ToString() + "]");
            if ((count < 0) || (start + count > tokensArray.Length))
                throw new ArgumentException("Invalid count value [" + start.ToString() + ", " + count.ToString() + "]");
            var tokensOut = new List<IToken>();
            for (int index = start; index < start + count; index++)
                tokensOut.Add(tokensArray[index]);
            return tokensOut.ToArray();
        }

        /// <summary>
        /// Return a new list that is a subset of the input token list - taken from the
        /// start position to the end of the token list
        /// </summary>
        protected IEnumerable<IToken> getTokenListSection(IEnumerable<IToken> tokens, int start)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            var tokensArray = tokens.ToArray();
            return getTokenListSection(tokensArray, start, tokensArray.Length - start);
        }
    }
}
