using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.ContentBreaking
{
    public static class StringBreaker
    {
        private static char[] WhiteSpaceCharsExceptLineReturn
            = Enumerable.Range((int)char.MinValue, (int)char.MaxValue)
                .Select(v => (char)v)
                .Where(c => (c != '\n') && char.IsWhiteSpace(c))
                .ToArray();

        /// <summary>
        /// Break down scriptContent into a combination of StringToken, CommentToken, UnprocessedContentToken and EndOfStatementNewLine instances (the
        /// end of statement tokens will not have been comprehensively handled).  This will never return null nor a set containing any null references.
        /// </summary>
        public static IEnumerable<IToken> SegmentString(string scriptContent)
        {
            if (scriptContent == null)
                throw new ArgumentNullException("scriptContent");

            var index = 0;
            var tokenContent = "";
            var tokens = new List<IToken>();
            while (index < scriptContent.Length)
            {
                var chr = scriptContent.Substring(index, 1);

                // Check for comment
                bool isComment;
                if (chr == "'")
                    isComment = true;
                else if (index <= (scriptContent.Length - 3))
                {
                    var threeChars = scriptContent.Substring(index, 3);
                    var fourthChar = (index == scriptContent.Length - 3) ? (char?)null : scriptContent[index + 3];
                    if (threeChars.Equals("REM", StringComparison.InvariantCultureIgnoreCase)
                    && ((fourthChar == null) || WhiteSpaceCharsExceptLineReturn.Contains(fourthChar.Value)))
                    {
                        isComment = true;
                        index += 2;
                    }
                    else
                        isComment = false;
                }
                else
                    isComment = false;
                if (isComment)
                {
                    // Store any previous token content
                    bool isInlineComment;
                    if (tokenContent != "")
                    {
                        isInlineComment = !tokenContent.Contains('\n');
                        tokens.Add(new UnprocessedContentToken(tokenContent));
                        tokenContent = "";
                    }
                    else
                        isInlineComment = false;

                    // Move past comment marker and look for end of comment (end of the
                    // line) then store in a CommentToken instance
                    // - Note: Always want an EndOfStatementNewLineToken to appear before
                    //   comments, so ensure this is the case (if the previous token was
                    //   a Comment it doesn't matter, if the previous statement was a
                    //   String we'll definitely need an end-of-statement, if the
                    //   previous was Unprocessed, we only need end-of-statement
                    //   if the content didn't end with a line-return)
                    index++;
                    int breakPoint = scriptContent.IndexOf("\n", index);
                    if (breakPoint == -1)
                        breakPoint = scriptContent.Length;
                    if (tokens.Count > 0)
                    {
                        var prevToken = tokens[tokens.Count - 1];
                        if (prevToken is StringToken)
                        {
                            // StringToken CAN'T contain end-of-statement content so
                            // we'll definitely need an EndOfStatementNewLineToken
                            tokens.Add(new EndOfStatementSameLineToken());
                        }
                        else if (prevToken is UnprocessedContentToken)
                        {
                            // UnprocessedContentToken MAY conclude with end-of-statement
                            // content, we'll need to check
                            if (!prevToken.Content.TrimEnd(WhiteSpaceCharsExceptLineReturn).EndsWith("\n"))
                            {
                                tokens.RemoveAt(tokens.Count - 1);
                                tokens.Add(new UnprocessedContentToken(
                                    prevToken.Content.TrimEnd('\t', '\r', ' ')
                                ));
                                tokens.Add(new EndOfStatementSameLineToken());
                            }
                        }
                    }
                    var commentContent = scriptContent.Substring(index, breakPoint - index);
                    if (isInlineComment)
                        tokens.Add(new InlineCommentToken(commentContent));
                    else
                        tokens.Add(new CommentToken(commentContent));
                    index = breakPoint;
                }

                // Check for string content
                else if (chr == "\"")
                {
                    // Store any previous token content
                    if (tokenContent != "")
                    {
                        tokens.Add(new UnprocessedContentToken(tokenContent));
                        tokenContent = "";
                    }

                    // Try to grab string content
                    var indexString = index + 1;
                    while (true)
                    {
                        chr = scriptContent.Substring(indexString, 1);
                        if (chr == "\n")
                            throw new Exception("Encountered line return in string content");
                        if (chr != "\"")
                            tokenContent += chr;
                        else
                        {
                            // Quote character - is it doubled (ie. escaped quote)?
                            string chrNext;
                            if (indexString < (scriptContent.Length - 1))
                                chrNext = scriptContent.Substring(indexString + 1, 1);
                            else
                                chrNext = null;
                            if (chrNext == "\"")
                            {
                                // Escaped quote: push past and add singe chr to content
                                indexString++;
                                tokenContent += "\"";
                            }
                            else
                            {
                                // Non-escaped quote: string end
                                tokens.Add(new StringToken(tokenContent));
                                tokenContent = "";
                                index = indexString;
                                break;
                            }
                        }
                        indexString++;
                    }
                }

                // Check for crazy VBScript escaped-name variable content
                // - It's acceptable to name a variable pretty much anything if it's wrapped in square brackets; seems to be any character other than
                //   line returns and a closing square bracket (since there is no support for escaping the closing bracket). This includes single and
                //   double quotes, whitespace, colons, numbers, underscores, anything - in fact a valid variable name is [ ], meaning a single space
                //   wrapped in square brackets! This is a little-known feature but it shouldn't be too hard to parse out at this point.
                else if (chr == "[")
                {
                    // Store any previous token content
                    if (tokenContent != "")
                        tokens.Add(new UnprocessedContentToken(tokenContent));

                    tokenContent = "[";
                    var indexString = index + 1;
                    while (true)
                    {
                        chr = scriptContent.Substring(indexString, 1);
                        if (chr == "\n")
                            throw new Exception("Encountered line return in escaped-content variable name");
                        tokenContent += chr;
                        if (chr == "]")
                        {
                            tokens.Add(AtomToken.GetNewToken(tokenContent));
                            tokenContent = "";
                            index = indexString;
                            break;
                        }
                        indexString++;
                    }
                }

                // Mustn't be neither comment, string nor VBScript-escaped-variable-name..
                else
                    tokenContent += chr;

                // Move to next character (if any)..
                index++;
            }

            // Don't let any unhandled content get away!
            if (tokenContent != "")
                tokens.Add(new UnprocessedContentToken(tokenContent));

            return tokens;
        }
    }
}
