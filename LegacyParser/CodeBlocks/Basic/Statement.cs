using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class Statement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// A statement should be solely an expression with no return value or whose return value is ignored - eg. a function call. Tokens that
        /// represent a statement where the return value is used to set another variable's value should be described by a ValueSettingStatement.
        /// It is recommended to pass any tokens that are thought to be one Statement through the StatementHandler, which will break down the
        /// content into multiple Statements (if there are any AbstractEndOfStatementToken tokens).
        /// </summary>
        public Statement(IEnumerable<IToken> tokens, CallPrefixOptions callPrefix)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (!Enum.IsDefined(typeof(CallPrefixOptions), callPrefix))
                throw new ArgumentOutOfRangeException("callPrefix");

            Tokens = tokens.ToList().AsReadOnly();
            if (!Tokens.Any())
                throw new ArgumentException("Statements must contain at least one token");
            if (!Tokens.Any())
                throw new ArgumentException("Empty tokens specified - invalid");
            if (Tokens.Any(t => t == null))
                throw new ArgumentException("Null token passed into Statement constructor");
            var firstTokenAsAtom = tokens.First() as AtomToken;
            if ((firstTokenAsAtom != null) && firstTokenAsAtom.Content.Equals("Call", StringComparison.InvariantCultureIgnoreCase))
                throw new ArgumentException("The first token may not be the Call keyword, that must be specified through the CallPrefixOption value where present");

            BracketStandardisedTokens = GetBracketStandardisedTokens(Tokens);
            CallPrefix = callPrefix;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null, empty nor contain any nulls. The first token will not be an AtomToken for the CALL keyword, since
        /// the presence of that in the source should be described by the CallPrefix value. The bracketing may be irregular around the
        /// arguments of the method call (if this is a method call with arguments) due to VBScript's flexibility on this matter. To
        /// get a standardised representation (with consistent bracketing), consult the BracketStandardisedTokens property.
        /// </summary>
        public IEnumerable<IToken> Tokens { get; private set; }

        /// <summary>
        /// VBScript allow flexibility with brackets for method calls whose return value is not considered - eg. "Test 1" is acceptable
        /// while "Test(1)" might be considered more consistent. The brackets are required by statements described by the class
        /// ValueSettingStatement - eg. "a = Test 1" is not valid while "a = Test(1)" is. This method will return a token
        /// stream based on the current Statement's code, but with the optional brackets inserted where absent.
        /// </summary>
        public IEnumerable<IToken> BracketStandardisedTokens { get; private set; }

        public CallPrefixOptions CallPrefix { get; private set; }

        public enum CallPrefixOptions
        {
            Absent,
            Present
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            var tokensList = Tokens.ToList();
            if (CallPrefix == CallPrefixOptions.Present)
                tokensList.Insert(0, AtomToken.GetNewToken("Call"));

            var output = new StringBuilder();
            output.Append(indenter.Indent);
            for (int index = 0; index < tokensList.Count; index++)
            {
                var token = tokensList[index];
                if (token is StringToken)
                    output.Append("\"" + token.Content + "\"");
                else
                    output.Append(token.Content);

                var nextToken = (index < (tokensList.Count - 1)) ? tokensList[index + 1] : null;
                if (nextToken == null)
                    continue;

                if ((token is MemberAccessorOrDecimalPointToken)
                || (token is OpenBrace)
                || (nextToken is MemberAccessorOrDecimalPointToken)
                || (nextToken is ArgumentSeparatorToken)
                || (nextToken is OpenBrace)
                || (nextToken is CloseBrace))
                    continue;

                output.Append(" ");
            }
            return output.ToString().TrimEnd();
        }

        // =======================================================================================
        // STANDARDISE BRACKETS, REMOVE VBScript WOBBLINESS
        //   "Test 1" bad
        //   "Test(1) good
        //   "a.Test 1" bad
        //   "a.Test(1)" good
        //   "a(0).Test 1" bad
        //   "a(0).Test(1)" good
        // Note: Allow "Test" or "a.Test", if there are no arguments then it's less important
        // =======================================================================================
        private static IEnumerable<IToken> GetBracketStandardisedTokens(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            var tokenArray = tokens.ToArray();
            if (tokenArray.Length == 0)
                throw new ArgumentException("Empty tokens set specified - invalid for a Statement");

            var bracketCount = 0;
            IToken lastUnbracketedToken = null;
            for (var tokenIndex = 0; tokenIndex < tokenArray.Length; tokenIndex++)
            {
                var token = tokenArray[tokenIndex];
                if (tokenIndex == 0)
                {
                    // The most common forms will be to start with a base AtomToken or KeyWordToken but the statement "Call(a.Test(1))" would appear
                    // here as "(a.Test(1))" as the "Call" keyword is removed and so an OpenBrace is a valid first token as well
                    if (token is OpenBrace)
                        bracketCount = 1;
					else if (!IsTokenAcceptableToCommenceCallExecution(token))
                        throw new ArgumentException("The first token should be an AtomToken or a KeyWordToken (not another type derived from AtomToken) to be a valid Statement");
                }
                else
                {
                    // Keep track of any bracketed content, this may be the arguments of a method call, meaning that the brackets are already "standard",
                    // or it may be before the arguments - eg. "a(1, 0).Test 1". Bracketed content can be ignored, we need to keep track of when we hit
                    // content that isn't bracketed, that should be.
                    if (token is OpenBrace)
                        bracketCount++;
                    else if (token is CloseBrace)
                    {
                        if (bracketCount == 0)
                            throw new ArgumentException("Invalid input, mismatched brackets (encountered a close bracket with no corresponding opening)");
                        bracketCount--;
                    }
                    if (bracketCount == 0)
                    {
                        var insertBracketsBeforeThisToken = false;
                        if (token is StringToken)
                        {
                            // If we've reached an un-bracketed string then we need to standardise the brackets starting before this token and closing around the last
                            insertBracketsBeforeThisToken = true;
                        }
						if (IsTokenAcceptableToCommenceCallExecution(token)
                        && (lastUnbracketedToken != null)
						&& (IsTokenAcceptableToCommenceCallExecution(lastUnbracketedToken)))
                        {
                            // If we've hit adjacent tokens (excluding bracketed content) that look like objects, properties or functions then there should
                            // be brackets in between. This covers cases such as
                            //   "Test 1" bad
                            //   "a.Test 1" bad
                            //   "a(0).Test 1" bad
                            insertBracketsBeforeThisToken = true;
                        }
                        if (insertBracketsBeforeThisToken)
                        {
                            return tokenArray.Take(tokenIndex)
                                .Concat(new[] { new OpenBrace("(") })
                                .Concat(tokenArray.Skip(tokenIndex))
                                .Concat(new[] { new CloseBrace(")") });
                        }
                    }
                }
                if ((bracketCount == 0) && !(token is CloseBrace))
                    lastUnbracketedToken = token;
            }
            if (bracketCount > 0)
                throw new ArgumentException("Invalid input, mismatched brackets (open brackets were not all closed)");

            // If we've got all the way to here then the content must be fine as it is, no changes were required
            return tokenArray;
        }

		/// <summary>
		/// Is the current token something that could represent the start of, or the entirety of, a call or value access - values may be of types
		/// NumericValueToken or BuiltInValueToken (in which case they should be entirety of an expression segment) or it might be the start of
		/// a function call (eg.  BuiltInFunctionToken or NameToken, in some cases). Tokens that would not be acceptable would be open braces
		/// (since bracketed expressions should be handled separately above) or ArgumentSeparatorToken, amongst others.
		/// </summary>
        private static bool IsTokenAcceptableToCommenceCallExecution(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return (
                (token.GetType() == typeof(AtomToken)) ||

                (token.GetType() == typeof(BuiltInFunctionToken)) ||
                (token.GetType() == typeof(BuiltInValueToken)) ||
                (token.GetType() == typeof(KeyWordToken)) ||

                (token is NameToken) ||
                (token.GetType() == typeof(NumericValueToken)) ||
                (token.GetType() == typeof(StringToken))
            );
        }
    }
}
