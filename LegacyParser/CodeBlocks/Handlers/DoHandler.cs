using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class DoHandler : AbstractBlockHandler
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

            // Determine whether we've got a "DO", "DO WHILE" or "DO UNTIL"
            if (!base.checkAtomTokenPattern(tokens, new string[] { "DO" }, false))
                return null;
            if (tokens.Count < 3)
                throw new ArgumentException("Insufficient tokens - invalid");
            IToken token = base.getToken(tokens, 1, new List<Type> { typeof(AtomToken), typeof(AbstractEndOfStatementToken) });
            bool hasCondition, doUntil;
            if (token is AtomToken)
            {
                bool doWhile = (token.Content.ToUpper() == "WHILE");
                doUntil = (token.Content.ToUpper() == "UNTIL");
                if (!doWhile && !doUntil)
                    throw new Exception("Invalid content - not EndOfStatement, WHILE or UNTIL after DO");
                hasCondition = true;
            }
            else
            {
                hasCondition = false;
                doUntil = false;
            }

            // Remove DO keyword and grab conditional content (if any)
            tokens.RemoveAt(0);
            Expression conditionStatement;
            if (!hasCondition)
                conditionStatement = null;
            else
            {
                // Loop for end of line..
                tokens.RemoveAt(0); // Remove WHILE / UNTIL
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
                conditionStatement = new Expression(tokensInCondition);
            }

            // Get block content
            string[] endSequenceMet;
            List<string[]> endSequences = new List<string[]>()
            {
                new string[] { "LOOP" }
            };
            CodeBlockHandler codeBlockHandler = new CodeBlockHandler(endSequences);
            List<ICodeBlock> blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find end sequence!");

            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if ((tokens.Count > 0) && (!(tokens[0] is AbstractEndOfStatementToken)))
                throw new Exception("EndOfStatementToken missing after LOOP");
            else
                tokens.RemoveAt(0);

            // Return Function code block instance
            return new DoBlock(conditionStatement, doUntil, blockContent);
        }
    }
}
