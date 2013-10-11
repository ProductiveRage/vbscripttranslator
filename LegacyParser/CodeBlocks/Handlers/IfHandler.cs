using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    /// <summary>
    /// This was one of the most complicated block handlers to implement! It could probably be tidied up some, but would need careful attention (and unit tests)
    /// </summary>
    public class IfHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            // Input validation
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;
            
            // Can we handle this point in the token stream?
            IToken firstToken = tokens[0];
            if ((!(firstToken is AtomToken)) || (firstToken.Content.ToUpper() != "IF"))
                return null;

            // Find the matching "THEN" for the "IF"
            List<IToken> expression = new List<IToken>();
            int indexThen = -1;
            for (int index = 0; index < tokens.Count; index++)
            {
                IToken token = tokens[index];
                if ((token is AtomToken) && (token.Content.ToUpper() == "THEN"))
                {
                    indexThen = index;
                    break;
                }
                else
                    expression.Add(token);
            }
            if (indexThen == -1)
                throw new Exception("IF statement not got matching THEN");

            // Is next token an EndOfStatementToken or is it a one-line IF statement?
            if (indexThen == (tokens.Count - 1))
                throw new Exception("IF statement's THEN keyword is final token - missing content");
            if (tokens[indexThen + 1] is EndOfStatementNewLineToken)
                return processMultiLine(tokens, indexThen);
            return processSingleLine(tokens);
        }

        private ICodeBlock processMultiLine(List<IToken> tokens, int offsetToTHEN)
        {
            // ======================================================================
            // Notes on multi-line If's:
            // ======================================================================
            // - Invalid: Combining first statement with IF line
            //   eg. IF (1) THEN Func
            //        Func
            //       END IF
            //
            // - Invalid: Trailing same-line-end-of-statement after IF
            //   eg. IF (1) THEN :
            //        Func(0)
            //       END IF
            // 
            // - Valid: Combining first statement with ELSEIF / ELSE line
            //   eg. IF (1) THEN
            //        Func
            //       ELSEIF (1) THEN Func
            //       ELSEIF (1) THEN
            //        Func
            //       ELSEIF (1) THEN Func
            //        Func
            //       ELSE Func
            //       END IF
            //
            // - Valid: Trailing same-line-end-of-statement after ELSEIF / ELSE
            //   eg. IF (1) THEN
            //        Func(0)
            //       ELSEIF(1) THEN :
            //        Func(0)
            //       END IF
            // 
            // - Invalid: ELSE / ELSEIF must always be the first statement on the line
            //   eg. IF (0) THEN
            //        Func(1):Func(2):ELSE
            //       END IF
            //
            // - Valid: IF may be combined with a previous line
            //   eg. Func: IF (0) THEN
            //        Func:Func:ELSE
            //       END IF
            // ======================================================================
            // Grab content inside IF blocks
            List<string[]> endSequences = new List<string[]>()
            {
                new string[] { "END", "IF" },
                new string[] { "ELSEIF" },
                new string[] { "ELSE" }
            };
            string[] endSequenceMet = new string[] { "IF" };
            List<IfBlock.IfBlockSegment> ifContent = new List<IfBlock.IfBlockSegment>();
            List<IToken> conditionTokens = null;
            while (true)
            {
                // First loop, we'll match this as the initial "IF" statement so we can
                // grab the content after it..
                if (endSequenceMet == null)
                    throw new Exception("Didn't find end sequence!");
                
                // Remove the last end sequence tokens (on the first loop, this will
                // actually remove the initial "IF" token)
                tokens.RemoveRange(0, endSequenceMet.Length);

                // The sequences that indicate further content are single-token ("IF",
                // "ELSEIF", ELSE")
                if (endSequenceMet.Length == 1)
                {
                    // An "ELSE" token won't have any condition content
                    if (endSequenceMet[0].ToUpper() == "ELSE")
                        conditionTokens = null;
                    else
                    {
                        // Grab the condition content (up to the "THEN" token) - this will
                        // throw an exception if it can't find "THEN" or there is no content
                        conditionTokens = getConditionContent(tokens).ToList();
                    }

                    // If the first character is a new-line end-of-statement (which is a
                    // common form), we may as well pull that out as well
                    if ((tokens.Count > 0) && (tokens[0] is EndOfStatementNewLineToken))
                        tokens.RemoveAt(0);
                }

                // Try to find the end of this content block (this will remove tokens
                // from the stream that have been processed by the block handler)
                var codeBlockHandler = new CodeBlockHandler(endSequences);
                var blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
                if (endSequenceMet == null)
                    throw new Exception("Didn't find end sequence!");

                // Record this content block
                if (conditionTokens == null)
                    ifContent.Add(new IfBlock.IfBlockElseSegment(blockContent));
                else
                {
                    Expression conditionStatement = new Expression(conditionTokens);
                    ifContent.Add(new IfBlock.IfBlockConditionSegment(conditionStatement, blockContent));
                }

                // If we hit, "END IF" then we're all done
                if ((endSequenceMet.Length == 2) && (endSequenceMet[0].ToUpper() == "END"))
                {
                    // Need to remove the final tokens - if we don't break out of the loop
                    // here, then end sequence tokens are removed up at the top
                    tokens.RemoveRange(0, endSequenceMet.Length);
                    break;
                }
            }

            // Expect the next token will be end-of-statement for IF block - drop that too
            if ((tokens.Count > 0) && (tokens[0] is AbstractEndOfStatementToken))
                tokens.RemoveAt(0);

            // Return fully-formed IfBlock
            return new IfBlock(ifContent);
        }

        private ICodeBlock processSingleLine(List<IToken> tokens)
        {
            // ======================================================================
            // Notes on single-line If's:
            // ======================================================================
            // - Invalid: Combining single line with multi-line
            //   eg. IF (1) THEN Func() ELSE IF (1) THEN
            //        Func()
            //       END IF
            //
            // - Invalid: ElseIf formations
            //   eg. IF (1) THEN Func() ELSEIF (2) THEN Func()
            //
            // - Combined If statements ARE valid
            //   eg. IF (1) THEN Func() ELSE IF (2) THEN Func()
            //   eg. IF (1) THEN IF (2) THEN Func()
            //
            // - Empty initial statement on If, with same-line end-of-statement
            //   eg. IF (1) THEN : Func(2)
            //
            // - Dropping statement from else is valid
            //   eg. IF (1) THEN Func() ELSE
            //
            // - Nested if statements can only ELSE to the last statement
            //   Valid:   IF (1) THEN Func() ELSE IF (1) THEN Func() ELSE Func()
            //   InValid: IF (1) THEN Func() ELSE IF (1) THEN Func() ELSE Func() ELSE Func()
            // ======================================================================
            // Grab content up to new line token (or end of content)
            List<IToken> ifTokens = new List<IToken>();
            foreach (IToken token in tokens)
            {
                if (token is EndOfStatementNewLineToken)
                    break;
                else
                {
                    if ((!(token is AtomToken))
                    && (!(token is StringToken))
                    && (!(token is EndOfStatementSameLineToken)))
                        throw new Exception("IfHandler.processSingleLine: Encountered invalid Token - should all be AtomToken, StringToken or EndOfStatementSameLineToken until new-line end of statement");
                    ifTokens.Add(token);
                }
            }

            // We'll need to store the number of tokens associated with this IF statement
            // because we'll pull them out of the stream later (but we'll be manipulating
            // this ifTokens list as well, so stash the original count now)
            int ifTokensCount = ifTokens.Count;

            // Pull "IF" token from the start
            ifTokens.RemoveAt(0);

            // Look for "THEN" token (pull handled tokens out of stream) - this will
            // throw an exception if it can't find "THEN" or there is no content
            var conditionTokens = getConditionContent(ifTokens).ToList();

            // Look for "ELSE" token (if present), ensuring we don't encounter any
            // "ELSEIF" tokens (which aren't valid in a single line "IF")
            int offsetElse = -1;
            for (int index = 0; index < ifTokens.Count; index++)
            {
                IToken token = ifTokens[index];
                if (token is AtomToken)
                {
                    if (token.Content.ToUpper() == "ELSE")
                    {
                        offsetElse = index;
                        break;
                    }
                    else if (token.Content.ToUpper() == "ELSEIF")
                        throw new Exception("Invalid content: ELSEIF found in single-line IF");
                }
            }

            // We have the condition statement, now get the Post-THEN content and
            // Post-ELSE content (may be null) - ie. the condition's "met" and
            // "not met" code blocks
            List<IToken> truthTokens, notTokens;
            if (offsetElse == -1)
            {
                truthTokens = base.getTokenListSection(ifTokens, 0).ToList();
                notTokens = null;
            }
            else
            {
                truthTokens = base.getTokenListSection(ifTokens, 0, offsetElse).ToList();
                notTokens = base.getTokenListSection(ifTokens, offsetElse + 1).ToList();
            }

            // Note: It's not valid for Post-THEN content to be empty
            if (truthTokens.Count == 0)
                throw new Exception("Empty THEN content in IF");

            // If we've got this far, we can pull out the processed tokens from the input list
            // - If statement ended at new-line-end-of-statement (as opposed to end of
            //   token stream) then remove that token as well
            tokens.RemoveRange(0, ifTokensCount);
            if (tokens.Count > 0)
                tokens.RemoveAt(0);

            // Translate token sets into code blocks
            string[] endSequenceMet;
            CodeBlockHandler codeBlockHandler = new CodeBlockHandler(null);
            List<ICodeBlock> truthStatement = codeBlockHandler.Process(truthTokens, out endSequenceMet);
            List<ICodeBlock> notStatement;
            if (notTokens == null)
                notStatement = null;
            else
                notStatement = codeBlockHandler.Process(notTokens, out endSequenceMet);
            Expression conditionStatement = new Expression(conditionTokens);

            // Generate IfBlock
            List<IfBlock.IfBlockSegment> ifContent = new List<IfBlock.IfBlockSegment>();
            ifContent.Add(new IfBlock.IfBlockConditionSegment(conditionStatement, truthStatement));
            if (notStatement != null)
                ifContent.Add(new IfBlock.IfBlockElseSegment(notStatement));
            return new IfBlock(ifContent);
        }

        /// <summary>
        /// While processing an IF / ELSEIF statement - that keyword should have been removed
        /// from the token stream, such that the start of the stream is the start of the
        /// condition content. This will trim off and return the condition content. An
        /// exception will be thrown for invalid content, but neither null nor an
        /// empty list will ever be returned.
        /// </summary>
        private IEnumerable<IToken> getConditionContent(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("token");

            // Look for "THEN" token - only acceptable tokens are Atom or String
            int offsetThen = -1;
            for (var index = 0; index < tokens.Count; index++)
            {
                var token = tokens[index];
                if ((!(token is AtomToken)) && (!(token is StringToken)))
                    throw new Exception("Encountered invalid token looking for THEN content");
                if ((token is AtomToken) && (token.Content.ToUpper() == "THEN"))
                {
                    offsetThen = index;
                    break;
                }
            }
            if (offsetThen == -1)
                throw new Exception("Invalid content: no THEN token for IF / ELSEIF");

            // Grab condition content
            var conditionTokens = base.getTokenListSection(tokens, 0, offsetThen);

            // Trim out the handled tokens and the "THEN"
            tokens.RemoveRange(0, conditionTokens.Count() + 1);

            // Return the content
            return conditionTokens;
        }
    }
}
