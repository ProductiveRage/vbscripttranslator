using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class WhileHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// Note: This will return a DoBlock if it can process content, as the While structure is effecively just a restricted Do
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            // Determine whether we've got a "WHILE" block
            if (!base.checkAtomTokenPattern(tokens, new string[] { "WHILE" }, false))
                return null;
            if (tokens.Count < 3)
                throw new ArgumentException("Insufficient tokens - invalid");

            // Remove WHILE keyword and grab conditional content
            tokens.RemoveAt(0);

            // Loop for end of line..
            List<IToken> tokensInCondition = new List<IToken>();
            while (true)
            {
                // Add AtomTokens to list until find EndOfStatement
                if (base.isEndOfStatement(tokens, 0))
                {
                    tokens.RemoveAt(0);
                    break;
                }
                IToken tokenCondition = base.getToken_AtomOrStringOnly(tokens, 0);
                tokensInCondition.Add(tokenCondition);
                tokens.RemoveAt(0);
            }
            Expression conditionStatement = new Expression(tokensInCondition);

            // Get block content
            string[] endSequenceMet;
            List<string[]> endSequences = new List<string[]>()
            {
                new string[] { "WEND" }
            };
            CodeBlockHandler codeBlockHandler = new CodeBlockHandler(endSequences);
            List<ICodeBlock> blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find end sequence!");

            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if (tokens.Count > 0)
            {
                if (tokens[0] is AbstractEndOfStatementToken)
                    tokens.RemoveAt(0);
                else
                    throw new Exception("EndOfStatementToken missing after WHILE");
            }

            // Return code block instance
            return new DoBlock(conditionStatement, isPreCondition: true, doUntil: false, statements: blockContent);
        }
    }
}
