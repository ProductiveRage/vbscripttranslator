using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class ForHandler : AbstractBlockHandler
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

            // Determine whether we've got a "FOR" or "FOR EACH" block
            if (base.checkAtomTokenPattern(tokens, new string[] { "FOR", "EACH" }, false))
                return handleForEach(tokens);
            else if (base.checkAtomTokenPattern(tokens, new string[] { "FOR" }, false))
                return handleForStandard(tokens);
            return null;
        }

        private ICodeBlock handleForEach(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            if (!base.checkAtomTokenPattern(tokens, new string[] { "FOR", "EACH"}, false))
                throw new ArgumentException("Invalid tokens - doesn't start FOR EACH");
            if (!base.checkAtomTokenPattern(tokens, 3, new string[] { "IN"}, false))
                throw new ArgumentException("Invalid tokens - doesn't start have IN keyword");
            if (!(tokens[2] is AtomToken))
                throw new ArgumentException("Invalid content - variable name is not AtomToken");

            // Grab loop variable name
            var loopVarToken = tokens[2];

            // Grab loop "base" (the collection to be looped through)
            List<IToken> loopSrc = getExpressionContent(tokens, 4);

            // Removed process content (loopSrc tokens + FOR + EACH + loopVar + IN + end-of-statement)
            if (!base.isEndOfStatement(tokens, loopSrc.Count + 4))
                throw new ArgumentException("Invalid content - didn't encounter end-of-statement after FOR EACH declaration");
            tokens.RemoveRange(0, loopSrc.Count + 5);

            // Get block content
            var blockContent = getForBlockContent(tokens);
            return new ForEachBlock(
                new NameToken(loopVarToken.Content, loopVarToken.LineIndex),
                new Expression(loopSrc),
                blockContent
            );
        }
        
        private ICodeBlock handleForStandard(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            if (!base.checkAtomTokenPattern(tokens, new string[] { "FOR" }, false))
                throw new ArgumentException("Invalid tokens - doesn't start FOR EACH");
            if (!base.checkAtomTokenPattern(tokens, 2, new string[] { "=" }, false))
                throw new ArgumentException("Invalid tokens - doesn't start have \"=\" comparison");
            if (!(tokens[2] is AtomToken))
                throw new ArgumentException("Invalid content - variable name is not AtomToken");

            // Grab loop variable name
            var loopVarToken = tokens[1];

            // Grab "from" expression
            List<IToken> loopFrom = getExpressionContent(tokens, 3, "TO");

            // Grab "to" expression
            List<IToken> loopTo = getExpressionContent(tokens, 4 + loopFrom.Count, "STEP");

            // Ensure we hit either end-of-statement or "STEP"??
            if (tokens.Count < 4 + loopFrom.Count + loopTo.Count)
                throw new ArgumentException("Insufficient token content");
            IToken tokenNext = tokens[4 + loopFrom.Count + loopTo.Count];
            List<IToken> stepExpr;
            if (tokenNext is AbstractEndOfStatementToken)
                stepExpr = null;
            else
                stepExpr = getExpressionContent(tokens, 5 + loopFrom.Count + loopTo.Count);
            
            // Remove processed tokens then get block content
            tokens.RemoveRange(0, 1); // "FOR"
            tokens.RemoveRange(0, 1); // loopVar
            tokens.RemoveRange(0, 1); // "="
            tokens.RemoveRange(0, loopFrom.Count);
            tokens.RemoveRange(0, 1); // "TO"
            tokens.RemoveRange(0, loopTo.Count);
            if (stepExpr != null)
            {
                tokens.RemoveRange(0, 1); // "STEP"
                tokens.RemoveRange(0, stepExpr.Count);
            }
            tokens.RemoveRange(0, 1); // End-of-statement
            var blockContent = getForBlockContent(tokens);
            
            // All done!
            return new ForBlock(
                new NameToken(loopVarToken.Content, loopVarToken.LineIndex),
                new Expression(loopFrom),
                new Expression(loopTo),
                (stepExpr == null ? null : new Expression(stepExpr)),
                blockContent
            );
        }

        private List<IToken> getExpressionContent(List<IToken> tokens, int offset, string endMarkerContent)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if ((offset < 0) || (offset >= tokens.Count))
                throw new ArgumentException("Invalid offset value [" + offset.ToString() + "]");
            if ((endMarkerContent != null) && (endMarkerContent.Trim() == ""))
                throw new ArgumentException("Blank endMarkerContent value - null is acceptable, blank is not");
            List<IToken> exprTokens = new List<IToken>();
            for (int index = offset; index < tokens.Count; index++)
            {
                if (tokens[index] is AbstractEndOfStatementToken)
                    break;
                IToken token = base.getToken_AtomOrStringOnly(tokens, index);
                if ((token is AtomToken) && (endMarkerContent != null))
                {
                    if (((AtomToken)token).Content.ToUpper() == endMarkerContent.ToUpper())
                        break;
                }
                exprTokens.Add(token);
            }
            return exprTokens;
        }
        private List<IToken> getExpressionContent(List<IToken> tokens, int offset)
        {
            return getExpressionContent(tokens, offset, null);
        }


        /// <summary>
        /// Note: The "FOR.." OR "FOR EACH".. content should have been removed from the
        /// token stream before calling this method
        /// </summary>
        private List<ICodeBlock> getForBlockContent(List<IToken> tokens)
        {
            string[] endSequenceMet;
            List<string[]> endSequences = new List<string[]>()
            {
                new string[] { "NEXT" }
            };
            CodeBlockHandler codeBlockHandler = new CodeBlockHandler(endSequences);
            List<ICodeBlock> blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find end sequence!");

            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if ((tokens.Count > 0) && (!(tokens[0] is AbstractEndOfStatementToken)))
                throw new Exception("EndOfStatementToken missing after NEXT");
            else
                tokens.RemoveAt(0);

            // Return code block instance
            return blockContent;
        }
    }
}
