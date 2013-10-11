using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class SelectHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            if (!base.checkAtomTokenPattern(tokens, new string[] { "SELECT", "CASE" }, false))
                return null;

            // Trim out "SELECT CASE" tokens
            tokens.RemoveRange(0, 2);
            
            // Grab content for the case expression
            List<IToken> expressionTokens = new List<IToken>();
            for (int index = 0; index < tokens.Count; index++)
            {
                if (base.isEndOfStatement(tokens, index))
                {
                    // Remove expression tokens (plus end-of-statement) from stream
                    tokens.RemoveRange(0, expressionTokens.Count + 1);
                    break;
                }

                // Add token to expression (must be Atom or String)
                expressionTokens.Add(base.getToken_AtomOrStringOnly(tokens, index));
            }

            // Look for the first CASE entry (note: it's allowable for there to be no
            // CASE entries at all, and case entries can be empty). It is also valid
            // to have comments outside of the CASE entries, though no other tokens
            // are valid in those areas.
            List<CommentStatement> openingComments = new List<CommentStatement>();
            List<IToken> tokensIgnored = new List<IToken>();
            for (int index = 0; index < tokens.Count; index++)
            {
                IToken token = tokens[index];
                if (token is CommentToken)
                    openingComments.Add(new CommentStatement(token.Content));
                else if (token is AbstractEndOfStatementToken)
                {
                    // Ignore blank lines
                    tokensIgnored.Add(token);
                }
                else if (token is AtomToken)
                {
                    if (token.Content.ToUpper() == "CASE")
                        break;
                    else if (token.Content.ToUpper() == "END")
                    {
                        if (index == (tokens.Count - 1))
                            throw new Exception("Error processing SELECT CASE block - reached end of token stream");
                        IToken tokenNext = tokens[index + 1];
                        if (!(tokenNext is AtomToken))
                            throw new Exception("Error processing SELECT CASE block - reached END followed invalid token [" + tokenNext.GetType().ToString() + "]");
                        if (tokenNext.Content.ToUpper() != "SELECT")
                            throw new Exception("Error processing SELECT CASE block - reached non-SELECT END tokens");
                        break;
                    }
                }
                else
                    throw new Exception("Invalid token encountered in SELECT CASE block [" + token.GetType().ToString() + "]");
            }
            tokens.RemoveRange(0, openingComments.Count + tokensIgnored.Count);

            // Unless we hit "END SELECT" straight away, process CASE blocks
            List<SelectBlock.CaseBlockSegment> content = new List<SelectBlock.CaseBlockSegment>();
            if (tokens[0].Content.ToUpper() != "END")
            {
                string[] endSequenceMet;
                List<string[]> endSequences = new List<string[]>()
                {
                    new string[] { "CASE" },
                    new string[] { "END", "SELECT" }
                };
                CodeBlockHandler codeBlockHandler = new CodeBlockHandler(endSequences);
                while (true)
                {
                    // Try to grab value(s) for CASE block
                    // - Get lists of tokens (may be multiple values, may be ELSE..)
                    List<List<IToken>> exprValues = base.getEntryList(tokens, 1, new EndOfStatementNewLineToken());
                    
                    // - Remove the CASE token
                    tokens.RemoveRange(0, 1);
                    // - Remove the exprValues tokens
                    foreach (List<IToken> valueTokens in exprValues)
                        tokens.RemoveRange(0, valueTokens.Count);
                    // - Remove the commas between expressions
                    tokens.RemoveRange(0, exprValues.Count - 1);
                    // - Remove the end-of-statement token
                    tokens.RemoveRange(0, 1);

                    // Quick check that it appears valid
                    bool caseElse = false;
                    if (exprValues.Count == 0)
                        throw new Exception("CASE block with no comparison value");
                    else
                    {
                        IToken firstExprToken = exprValues[0][0];
                        if ((firstExprToken is AtomToken) && (firstExprToken.Content.ToUpper() == "ELSE"))
                        {
                            if ((exprValues.Count > 1) || (exprValues[0].Count != 1))
                                throw new Exception("Invalid CASE ELSE opening statement");
                            caseElse = true;
                        }
                    }

                    // Try to grab single CASE block content
                    List<ICodeBlock> blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
                    if (endSequenceMet == null)
                        throw new Exception("Didn't find end sequence!");

                    // Add to CASE block list
                    if (caseElse)
                        content.Add(new SelectBlock.CaseBlockElseSegment(blockContent));
                    else
                    {
                        List<Expression> values = new List<Expression>();
                        foreach (List<IToken> valueTokens in exprValues)
                            values.Add(new Expression(valueTokens));
                        content.Add(new SelectBlock.CaseBlockExpressionSegment(values, blockContent));
                    }

                    // If we hit END SELECT then break out of loop, otherwise
                    // go back round to get the next block
                    if (endSequenceMet.Length == 2)
                    {
                        tokens.RemoveRange(0, endSequenceMet.Length);
                        if (tokens.Count > 0)
                        {
                            if (!(tokens[0] is AbstractEndOfStatementToken))
                                throw new Exception("EndOfStatementToken missing after END FUNCTION");
                            else
                                tokens.RemoveAt(0);
                        }
                        break;
                    }

                }
            }

            // All done!
            return new SelectBlock(new Expression(expressionTokens), openingComments, content);
        }
    }
}
