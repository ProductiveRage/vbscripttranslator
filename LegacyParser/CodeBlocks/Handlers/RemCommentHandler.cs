using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class RemCommentHandler : AbstractBlockHandler
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

            if (!base.checkAtomTokenPattern(tokens, new string[] { "REM" }, false))
                return null;

            int tokensConsumed = 0;
            StringBuilder commentContent = null;
            foreach (IToken token in tokens)
            {
                if (token == null)
                    throw new ArgumentException("Encountered null token in stream");
                tokensConsumed++;
                if (token is EndOfStatementNewLineToken)
                    break;
                else
                {
                    if (commentContent == null)
                        commentContent = new StringBuilder();
                    else
                        commentContent.Append(" ");
                    commentContent.Append(token.Content);
                }
            }
            tokens.RemoveRange(0, tokensConsumed);
            if (commentContent == null)
                return new CommentStatement("");
            return new CommentStatement(commentContent.ToString());
        }
    }
}
