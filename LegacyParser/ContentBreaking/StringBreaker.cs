using System;
using System.Collections.Generic;
using System.Linq;
using CSharpSupport;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.ContentBreaking
{
    public static class StringBreaker
    {
        private static char[] _whiteSpaceCharsExceptLineReturn
            = Enumerable.Range((int)char.MinValue, (int)char.MaxValue)
                .Select(v => (char)v)
                .Where(c => (c != '\n') && char.IsWhiteSpace(c))
                .ToArray();

        /// <summary>
        /// We don't know what culture the translated program will be running in, so we can't do full date literal validation here (if we're running the
        /// translation process in English then #1 May 2015# will be a valid date, but if the translated program runs in French then it won't be). With
        /// VBScript, date literals will be parsed each time the script is read and a syntax error raised if any of them are invalid, before any of the
        /// script is executed - this means that the script may behave differently in different cultures. This means that we can not assume any culture
        /// here and so can't fully verify date literals if they have a month name in them. What we CAN do is check the validity of date literals that
        /// are composed solely of numeric segments and we can check for some invalid date literals that contain month names (eg. #May June# is invalid
        /// since there are TWO months and #42 May 99# is invalid since the numbers 42 and 99 can not be used in any combination to represent a valid
        /// date). We can do this by returning 1 for any month name. It won't catch invalid dates such as "30 Feb 2012" and it won't catch invalid month
        /// names, but it's better than nothing. This compromise means that runtime checks are required at the start of the translated program, ensuring
        /// that an exception is raised before any work is performed if there are any invalid date literals present, considering the runtime culture.
        /// </summary>
        private static DateParser _limitedDateParser = new DateParser(
            monthNameTranslator: monthName => 1,
            defaultYearOverride: 2015
        );

        /// <summary>
        /// Break down scriptContent into a combination of StringToken, CommentToken, UnprocessedContentToken and EndOfStatementNewLine instances (the
        /// end of statement tokens will not have been comprehensively handled).  This will never return null nor a set containing any null references.
        /// </summary>
        public static IEnumerable<IToken> SegmentString(string scriptContent)
        {
            if (scriptContent == null)
                throw new ArgumentNullException("scriptContent");

            // Normalise line returns
            scriptContent = scriptContent.Replace("\r\n", "\n").Replace('\r', '\n');

            var index = 0;
            var tokenContent = "";
            var tokens = new List<IToken>();
            var lineIndex = 0;
			var lineIndexForStartOfContent = 0;
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
                    && ((fourthChar == null) || _whiteSpaceCharsExceptLineReturn.Contains(fourthChar.Value)))
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
                        // If there has been any one the same line as this comment, then this is an inline comment
                        var contentAfterLastLineReturn = tokenContent.Split('\n').Last();
                        isInlineComment = (contentAfterLastLineReturn.Trim() != "");
                        tokens.Add(new UnprocessedContentToken(tokenContent, lineIndexForStartOfContent));
                        tokenContent = "";
                    }
                    else
                        isInlineComment = false;

                    // Move past comment marker and look for end of comment (end of the line) then store in a CommentToken instance
                    // - Note: Always want an EndOfStatementNewLineToken to appear before comments, so ensure this is the case (if the previous token was
                    //   a Comment it doesn't matter, if the previous statement was a String we'll definitely need an end-of-statement, if the previous
                    //   was Unprocessed, we only need end-of-statement if the content didn't end with a line-return)
                    lineIndexForStartOfContent = lineIndex;
                    index++;
                    int breakPoint = scriptContent.IndexOf("\n", index);
                    if (breakPoint == -1)
                        breakPoint = scriptContent.Length;
                    if (tokens.Count > 0)
                    {
                        var prevToken = tokens[tokens.Count - 1];
                        if (prevToken is UnprocessedContentToken)
                        {
                            // UnprocessedContentToken MAY conclude with end-of-statement content, we'll need to check
                            if (!prevToken.Content.TrimEnd(_whiteSpaceCharsExceptLineReturn).EndsWith("\n"))
                            {
                                tokens.RemoveAt(tokens.Count - 1);
                                var unprocessedContentToRecord = prevToken.Content.TrimEnd('\t', ' ');
                                if (unprocessedContentToRecord != "")
                                {
                                    tokens.Add(new UnprocessedContentToken(unprocessedContentToRecord, prevToken.LineIndex));
                                    tokens.Add(new EndOfStatementSameLineToken(prevToken.LineIndex));
                                }
                            }
                        }
                    }
                    if (tokens.Any() && ((tokens.Last() is DateLiteralToken) || (tokens.Last() is StringToken)))
                    {
                        // Quoted literals (ie. string or date) CAN'T contain end-of-statement content so we'll definitely need an EndOfStatementNewLineToken
                        // Note: This has to be done after the above work in case there was a literal token then some whitespace (which is removed above)
                        // then a Comment. If the work above wasn't done before this check then "prevToken" would not be a StringToken, it would be the
                        // whitespace - but that would be removed and then the literal would be arranged right next to the Comment, without an end-
                        // of-statement token between them!
                        tokens.Add(new EndOfStatementSameLineToken(lineIndexForStartOfContent));
                    }
                    var commentContent = scriptContent.Substring(index, breakPoint - index);
                    if (isInlineComment)
						tokens.Add(new InlineCommentToken(commentContent, lineIndexForStartOfContent));
                    else
						tokens.Add(new CommentToken(commentContent, lineIndexForStartOfContent));
                    index = breakPoint;
                    lineIndex++;
                    lineIndexForStartOfContent = lineIndex;
                }

                // Check for string content
                else if (chr == "\"")
                {
                    // Store any previous token content
                    if (tokenContent != "")
                    {
						tokens.Add(new UnprocessedContentToken(tokenContent, lineIndexForStartOfContent));
                        tokenContent = "";
                    }

                    // Try to grab string content
                    lineIndexForStartOfContent = lineIndex;
                    var indexString = index + 1;
                    while (true)
                    {
                        chr = scriptContent.Substring(indexString, 1);
                        if (chr == "\n")
                            throw new Exception("Encountered line return in string content around line " + (lineIndexForStartOfContent + 1));
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
								tokens.Add(new StringToken(tokenContent, lineIndexForStartOfContent));
                                tokenContent = "";
								lineIndexForStartOfContent = lineIndex;
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
                        tokens.Add(new UnprocessedContentToken(tokenContent, lineIndexForStartOfContent));

                    lineIndexForStartOfContent = lineIndex;
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
                            tokens.Add(AtomToken.GetNewToken(tokenContent, lineIndexForStartOfContent));
                            tokenContent = "";
                            lineIndexForStartOfContent = lineIndex;
                            index = indexString;
                            break;
                        }
                        indexString++;
                    }
                }

                // VBScript supports date literals, wrapped in hashes. These introduce a range of complications - such as literal comparisons requiring
                // special logic, as string and number literals do - eg. ("a" = #2015-5-27#) will fail at runtime as "a" must be parse-able as a date,
                // and it isn't. It also has complications around culture - so the value #1 5 2015# must be parsed as 2015-5-1 in the UK when the
                // translated output is executed but as 2015-1-5 in the US. On top of that, VBScript is very flexible in its acceptance of date formats -
                // amongst these problems is that the year is optional and so #1 5# means 1st of May or 5th of January (depending upon culture) in the
                // current year - however, once a date literal has had a default year set for a given request it must stick to that year; so if the request
                // is unfortunate enough to be slow and cross years, a given date literal must consistently stick to using the year from when the request
                // started. When a new request starts, however, if the year has changed then that new request must default to that new year, it would be no
                // good if the year was determined once (at translation time) and then never changed, since this would be inconsistent with VBScript's behaviour
                // of treating each request as a whole new start-up / serve / tear-down process. This means that the value #29 2# will change by year, being
                // the 29th of February if the current year is a leap year and the 1st of February 2029 if not (since #29 2# will be interpreted as year 29
                // and month 2 since 29 could not be a valid month - and then 29 will be treated as a two-digit year which must be bumped up to 2029). Also
                // note that even in the US #29 2# will be interpreted as the 29th of February (or 1st of February 2029) since there is no way to parse that
                // as a month-then-day format).
                // - Note: This gets the lowest priority in terms of wrapping characters, so [#1 1#] is a variable name and not something containing a
                //   date, likewise "#1 1#" is a string and nothing to do with a date. There are no escape characters. If the wrapped value can not
                //   possibly be valid then an exception will be raised at this point.
                else if (chr == "#")
                {
                    // Store any previous token content
                    if (tokenContent != "")
                        tokens.Add(new UnprocessedContentToken(tokenContent, lineIndexForStartOfContent));

                    lineIndexForStartOfContent = lineIndex;
                    tokenContent = "";
                    var indexString = index + 1;
                    while (true)
                    {
                        chr = scriptContent.Substring(indexString, 1);
                        if (chr == "\n")
                            throw new Exception("Encountered line return in date literal content");
                        if (chr == "#")
                        {
                            // We can only catch certain kinds of invalid date literal format here since some formats are culture-dependent (eg. "1 May 2010" is
                            // valid in English but not in French) and I don't want to assume that translated programs are running with the same culture as the
                            // translation process. The "limitedDateParser" can catch some invalid formats, which is better than nothing, but others will have
                            // to checked at runtime (see the notes around the instantiation of the limitedDateParser).
                            try
                            {
                                _limitedDateParser.Parse(tokenContent);
                            }
                            catch (Exception e)
                            {
                                throw new ArgumentException("Invalid date literal content encountered on line " + lineIndex + ": #" + tokenContent + "#", e);
                            }
                            tokens.Add(new DateLiteralToken(tokenContent, lineIndexForStartOfContent));
                            tokenContent = "";
                            lineIndexForStartOfContent = lineIndex;
                            index = indexString;
                            break;
                        }
                        else
                            tokenContent += chr;
                        indexString++;
                    }
                }

                // Mustn't be neither comment, string, date nor VBScript-escaped-variable-name..
                else
                    tokenContent += chr;

                // Move to next character (if any)..
                index++;
				if (chr == "\n")
					lineIndex++;
			}

            // Don't let any unhandled content get away!
            if (tokenContent != "")
				tokens.Add(new UnprocessedContentToken(tokenContent, lineIndexForStartOfContent));

            return tokens;
        }
    }
}
