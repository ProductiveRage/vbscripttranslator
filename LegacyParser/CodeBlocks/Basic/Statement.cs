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
        /// ValueSettingStatement - eg. "a = Test 1" is not valid while "a = Test(1)" is (and so in that case, this "standardised"
        /// bracket retrieval would not be used, it would be used by non-value-returning statements only, such as "Test" or "Test 1").
        /// This method will return a token stream based on the current Statement's code, but with the optional brackets inserted where
        /// absent.
        /// </summary>
        public IEnumerable<IToken> GetBracketStandardisedTokens()
        {
            // The "bracket-standardised" content is only required when there is no return type for a statement - in which case it must
            // be possible to apply this logic. However, there are some times where a statement is valid but where this processing will
            // fail. For example, the optional "LoopStep" of a ForBlock will fail if it's value is "-x" since an operator is not considered
            // acceptable for the first token of a statement (since, if this were a standalone statement; "-x"; then there would be a
            // compile time error). This method was previously a property which was pre-evalaated in the constructor, but that caused
            // issues with ForBlocks in that format. So now, instead, this work is only done when required and so shouldn't cause
            // problems where it's not needed.
            return GetBracketStandardisedTokens(Tokens, CallPrefix);
        }

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
                tokensList.Insert(0, AtomToken.GetNewToken("Call", tokensList[0].LineIndex));

            var output = new StringBuilder();
            output.Append(indenter.Indent);
            for (int index = 0; index < tokensList.Count; index++)
            {
                var token = tokensList[index];
                if (token is StringToken)
                    output.Append("\"" + token.Content + "\"");
                else if (token is DateLiteralToken)
                    output.Append("#" + token.Content + "#");
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
        // 2014-03-24 DWR: This has been tightened up since "Test a" and "Test(a)" are not the
        // same to VBScript, the latter actually associates the brackets with the "a" and not
        // the "Test" call, brackets around an argument to function mean that the argument
        // should be passed ByVal, not ByRef - if the argument would otherwise be ByRef.
        // This is why "Test(1, 2)" is not valid (the error "Cannot use parentheses when calling
        // a Sub" will be raised) since the parentheses are associated with the function call,
        // which is not acceptable when not considering the return value - eg. "a = Test(1, 2)"
        // IS acceptable since the return value IS being considered. As such, the following
        // transformations should be applied:
        //   "Test a1" => "Test(a1)" (argument not forced to be ByVal)
        //   "Test(a1)" => "Test((a1))" (argument forced to be ByVal)
        //   "r = Test a1" is invalid ("Expected end of statement")
        //   "r = Test(a1)" requires no transformation (ByVal not enforced)
        //   "r = Test((a1))" requires no transformation (ByVal IS enforced)
        //   "Test(a1, a2)" is invalid ("Cannot use parentheses when calling a Sub")
        //   "r = Test(a1, a2)" requires no transformation (ByVal not enforced)
        //   "r = Test((a1), a2)" requires no transformation (ByVal enforced for a1 but not a2)
        //   "CALL Test(a1, a2)" requires no transformation (ByVal not enforced)
        //   "CALL Test((a1), a2)" requires no transformation (ByVal enforced for a1 but not a2)
        // See http://blogs.msdn.com/b/ericlippert/archive/2003/09/15/52996.aspx
        // (In all of the above examples, the value-setting statements such as "r = Test(a1)"
        // would not be represented by this class, the ValueSettingStatment class would be used,
        // but the examples are included to illustrate variations on the cases that must be
        // dealt with),
        // =======================================================================================
        private static IEnumerable<IToken> GetBracketStandardisedTokens(IEnumerable<IToken> tokens, CallPrefixOptions callPrefix)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (!Enum.IsDefined(typeof(CallPrefixOptions), callPrefix))
                throw new ArgumentOutOfRangeException("callPrefix");

            var tokenArray = tokens.ToArray();
            if (tokenArray.Length == 0)
                throw new ArgumentException("Empty tokens set specified - invalid for a Statement");

            // TODO: Add a method that will detect all forms of invalid content first and require that this method be called before the Statement is
            // processed further (dealing with all run time or compile time errors). Such a method would enable the work to be performed under the
            // assumption that the content was valid VBScript, which would make things easier. (Should this method be in a translator and this data
            // be returned from a method GetBracketStandardisedTokensIfContentValid?)

            // No need to try to re-arrange things if this is a new-instance expression, no brackets are required
            if ((tokenArray[0] is KeyWordToken) && tokenArray[0].Content.Equals("new", StringComparison.InvariantCultureIgnoreCase))
                return tokenArray;

            // 2014-04-21 DWR: Require some preliminary work to deal with "Test -1" such that it will be parsed as "Test(-1)" rather than an attempted
            // subtraction of 1 from Test. If Test is not a function with a single argument then there will be a runtime error raised. The case where
            // brackets are absent is only possible where CALL is not used (since that would make brackets usage compulsory).
            if (callPrefix == CallPrefixOptions.Absent)
            {
                var valueTokenTypes = new[] {
                    typeof(BuiltInFunctionToken),
                    typeof(BuiltInValueToken),
                    typeof(DateLiteralToken),
                    typeof(KeyWordToken),
                    typeof(NameToken),
                    typeof(NumericValueToken),
                    typeof(StringToken)
                };
                var firstItemTokens = new List<IToken>
                {
                    tokenArray[0]
                };
                for (var tokenIndex = 1; tokenIndex < tokenArray.Length; tokenIndex++)
                {
                    var token = tokenArray[tokenIndex];
                    if (token is MemberAccessorOrDecimalPointToken)
                    {
                        // If we assume valid content (which should be reasonable for the purposes of trying to pick up on this special case), then a
                        // MemberAccessorOrDecimalPointToken must indicate a continuation of the current term
                        firstItemTokens.Add(token);
                        continue;
                    }
                    if (valueTokenTypes.Contains(token.GetType()))
                    {
                        if (firstItemTokens.Last() is MemberAccessorOrDecimalPointToken)
                        {
                            // If the previous token was a "." and the content is valid (which we're assuming), then this token should be part of group
                            // containing the preceding tokens (eg. "WScript", ".", "Echo" are part of "WScript.Echo") so treat this as a further
                            // continuation of the current term
                            firstItemTokens.Add(token);
                            continue;
                        }
                        if (valueTokenTypes.Contains(firstItemTokens.Last().GetType()))
                        {
                            // If there are two adjacent value type tokens then they must be part of two terms - eg. "Test a" (if they were part of
                            // the same term then there would be a "joining" token such as a MemberAccessorOrDecimalPointToken - eg. "Test.a"). If
                            // we have reached this point without encountering a minus sign then the special case does not apply to this content.
                            break;
                        }
                    }

                    // If an open bracket is encountered then it must be bracketing an argument meaning that it will be treated as an "enforced ByVal"
                    // and so ADDITIONAL brackets are going to be required so that this information is not lost in the "standardised bracket" data. If
                    // a minus sign is encountered then it must be part of the first argument in a method call (until this point we wouldn't have known
                    // if it was an subtraction operation between two terms or if it was this case, where the minus sign is a negation of the second
                    // term, which is passed as an argument to the first term).
                    var isMinusSign = (token is OperatorToken) && (token.Content == "-");
                    if (isMinusSign)
                    {
                        // Note: Having hit a minus sign here before an open brace means that the brackets must be absent around the arguments - eg.
                        //   F1 -1
                        // We don't need to look ahead to see if the return value of this is considered, so we don't have to do any more. (If there
                        // are any more accessors to consider then they must be part of one of the arguments - eg.
                        //   F1 -1.1
                        // or
                        //   F1 -1, a.Name
                        // This is a less complicated case than it the current token is an OpenBrace (see below).
                        var remainingTokens = tokenArray.Skip(tokenIndex).ToArray();
                        if (isMinusSign && (remainingTokens.Length > 1))
                        {
                            // If this is a minus operator that starts the second term, but no brackets have been encountered, then insert brackets
                            // around the content that starts here (close the bracket at the end of the token set). VBScript gives special meaning to
                            // brackets in some cases - eg. "Test (a)" will mean that the argument is passed ByVal even when it would otherwise by ByRef,
                            // but the case that we're testing for here is where the first argument is made negative, which would mean it could never be
                            // passed ByRef anyway and so the change we do here can not have any functional impact.
                            var numericValue = remainingTokens[1] as NumericValueToken;
                            if (numericValue != null)
                            {
                                // If the first token in the to-be-bracketed content following the minus sign is a number then we can combine the
                                // minus operator and the number into a single token (this is similar to some of the work performed by the
                                // NumberRebuilder except that it doesn't analyse the code at this level and so could not pick up on this
                                // particular combination)
                                remainingTokens = new[] { numericValue.GetNegative() }
                                    .Concat(remainingTokens.Skip(2))
                                    .ToArray();
                            }
                        }
                        var lineIndexForInsertedOpenBrace = firstItemTokens.Last().LineIndex;
                        var lineIndexForInsertedCloseBrace = remainingTokens.Any() ? remainingTokens.Last().LineIndex : lineIndexForInsertedOpenBrace;
                        tokenArray = firstItemTokens
                            .Concat(new[] { new OpenBrace(lineIndexForInsertedOpenBrace) })
                            .Concat(remainingTokens)
                            .Concat(new[] { new CloseBrace(lineIndexForInsertedCloseBrace) })
                            .ToArray();
                        break;
                    }
                    if (token is OpenBrace)
                    {
                        // Note: Hitting an OpenBrace here may mean that we've hit the only argument content for a function and that it is enclosed in
                        // braces - eg.
                        //   F1(a)
                        // in which case F1 would be called with argument "a" passed ByVal, since the brackets in this non-value-returning function call
                        // would be given the special ByVal meaning. However, it is not conclusive, since we may actually be in the middle of content
                        // such as
                        //   F1(a).Name
                        // in which case the "F1(a)" IS a value-returning function call and the brackets do NOT have special significance in terms of the
                        // argument "a" being passed ByVal. In order to determine this, we need to look ahead in the content and count brackets in and
                        // out until we get to the end of the content (not necessarily the last token since the statement could be followed by a comment,
                        // for example). Until we get to that point we won't know what these brackets signify.
                        var remainingTokens = tokenArray.Skip(tokenIndex).ToArray();
                        var bracketedContent = new List<IToken> { remainingTokens.First() };
                        var bracketDepth = 1;
                        for (var index = 1; index < remainingTokens.Length; index++)
                        {
                            var bracketedContentToken = remainingTokens[index];
                            bracketedContent.Add(bracketedContentToken);
                            if (bracketedContentToken is CloseBrace)
                            {
                                bracketDepth--;
                                if (bracketDepth == 0)
                                    break;
                            }
                            else if (bracketedContentToken is OpenBrace)
                                bracketDepth++;
                        }
                        var lineIndexForInsertedOpenBrace = firstItemTokens.Last().LineIndex;
                        var lineIndexForInsertedCloseBrace = remainingTokens.Any() ? remainingTokens.Last().LineIndex : lineIndexForInsertedOpenBrace;
                        if (bracketDepth != 0)
                            throw new Exception("Invalid content - mismatched brackets ending on line " + (lineIndexForInsertedCloseBrace + 1));
                        var remainingTokensAfterBracketedContent = remainingTokens.Skip(bracketedContent.Count);
                        if (remainingTokensAfterBracketedContent.Any() && (remainingTokensAfterBracketedContent.First() is MemberAccessorOrDecimalPointToken))
                        {
                            // If there is a member accessor (a ".") after the bracketed content, then the brackets were a necessary part of a value-returning
                            // function call and not the "optional" brackets which mean that an argument was being forced to be passed ByVal.
                            firstItemTokens.AddRange(bracketedContent);
                            tokenIndex = firstItemTokens.Count - 1;
                            continue;
                        }
                        tokenArray = firstItemTokens
                            .Concat(new[] { new OpenBrace(lineIndexForInsertedOpenBrace) })
                            .Concat(bracketedContent)
                            .Concat(remainingTokensAfterBracketedContent)
                            .Concat(new[] { new CloseBrace(lineIndexForInsertedCloseBrace) })
                            .ToArray();
                        break;
                    }

                    // Any other token types encountered here (an operator, for example) mean that the special cases were not encountered and we can
                    // drop out of this loop
                    break;
                }
            }

            var bracketCount = 0;
            IToken lastUnbracketedToken = null;
            for (var tokenIndex = 0; tokenIndex < tokenArray.Length; tokenIndex++)
            {
                var nextTokenIfAny = (tokenArray.Length > 1) ? tokenArray[1] : null;
                var token = tokenArray[tokenIndex];
                if (tokenIndex == 0)
                {
                    // The most common forms will be to start with a base AtomToken or KeyWordToken but the statement "Call(a.Test(1))" would appear
                    // here as "(a.Test(1))" as the "Call" keyword is removed and so an OpenBrace is a valid first token as well
                    if (token is OpenBrace)
                        bracketCount = 1;
                    else if (!IsTokenAcceptableToCommenceCallExecution(token, nextTokenIfAny))
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
                        if ((token is DateLiteralToken) || (token is StringToken))
                        {
                            // If we've reached an un-bracketed string or date literal then we need to standardise the brackets starting before this token and closing around the last
                            insertBracketsBeforeThisToken = true;
                        }
						if (IsTokenAcceptableToCommenceCallExecution(token, nextTokenIfAny)
                        && (lastUnbracketedToken != null)
                        && (IsTokenAcceptableToCommenceCallExecution(lastUnbracketedToken, nextTokenIfAny)))
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
                                .Concat(new[] { new OpenBrace(tokenArray.Skip(tokenIndex).First().LineIndex) })
                                .Concat(tokenArray.Skip(tokenIndex))
                                .Concat(new[] { new CloseBrace(tokenArray.Last().LineIndex) });
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
        private static bool IsTokenAcceptableToCommenceCallExecution(IToken token, IToken nextTokenIfAny)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            // If the current token is a MemberAccessorOrDecimalPointToken and the next token is a NameToken then this must be a statement that
            // accesses a method or property within a "WITH" construct (in which case, the method or property accesses will need to be resolved
            // later on in the processing, according to the containing WITH target)
            if ((token is MemberAccessorOrDecimalPointToken) && (nextTokenIfAny != null) && (nextTokenIfAny is NameToken))
                return true;

            return (
                (token.GetType() == typeof(AtomToken)) ||

                (token.GetType() == typeof(BuiltInFunctionToken)) ||
                (token.GetType() == typeof(BuiltInValueToken)) ||
                (token.GetType() == typeof(KeyWordToken)) ||

                (token is NameToken) ||
                (token.GetType() == typeof(DateLiteralToken)) ||
                (token.GetType() == typeof(NumericValueToken)) ||
                (token.GetType() == typeof(StringToken))
            );
        }
    }
}
