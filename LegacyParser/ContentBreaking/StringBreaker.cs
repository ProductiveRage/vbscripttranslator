using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.ContentBreaking
{
    public static class StringBreaker
    {
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
                if (chr == "'")
                {
                    // Store any previous token content
                    if (tokenContent != "")
                    {
                        tokens.Add(new UnprocessedContentToken(tokenContent));
                        tokenContent = "";
                    }

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
                            tokens.Add(new EndOfStatementNewLineToken());
                        }
                        else if (prevToken is UnprocessedContentToken)
                        {
                            // UnprocessedContentToken MAY conclude with end-of-statement
                            // content, we'll need to check
                            if (!prevToken.Content.TrimEnd('\t', '\r', ' ').EndsWith("\n"))
                            {
                                tokens.RemoveAt(tokens.Count - 1);
                                tokens.Add(new UnprocessedContentToken(
                                    prevToken.Content.TrimEnd('\t', '\r', ' ')
                                ));
                                tokens.Add(new EndOfStatementNewLineToken());
                            }
                        }
                    }
                    tokens.Add(
                        new CommentToken(
                            scriptContent.Substring(index, breakPoint - index)
                        )
                    );
                    index = breakPoint;
                }

                // Check for string start / end
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

                // Must be neither comment nor string..
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
