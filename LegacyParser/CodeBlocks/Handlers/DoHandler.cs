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

            var preConditionStartToken = base.getToken(tokens, 1, new List<Type> { typeof(AtomToken), typeof(AbstractEndOfStatementToken) });
            bool hasPreCondition, doUntil;
            if (preConditionStartToken is AtomToken)
            {
                var doWhile = (preConditionStartToken.Content.ToUpper() == "WHILE");
                doUntil = (preConditionStartToken.Content.ToUpper() == "UNTIL");
                hasPreCondition = doWhile || doUntil;
            }
            else
            {
                hasPreCondition = false;
                doUntil = false;
            }

            // Remove DO keyword and grab pre-condition content (if any)
            tokens.RemoveAt(0);
            Expression conditionStatement;
            if (!hasPreCondition)
                conditionStatement = null;
            else
            {
                tokens.RemoveAt(0); // Remove WHILE / UNTIL
                conditionStatement = ExtractConditionFromTokens(tokens);
            }

            // If the token was a WHILE or UNTIL and we extracted a pre-condition following it, then the next token must be an end-of-statement. If there was
            // no pre-condition then we've only processed the "DO" and the next token must still be an end-of-statement.
            if (tokens.Count > 0)
            {
                if (tokens[0] is AbstractEndOfStatementToken)
                    tokens.RemoveAt(0);
                else
                    throw new ArgumentException("Expected end-of-statement token after DO keyword (relating to a construct without a pre-condition)");
            }

            // Get block content
            string[] endSequenceMet;
            var endSequences = new List<string[]>()
            {
                new string[] { "LOOP" }
            };
            var codeBlockHandler = new CodeBlockHandler(endSequences);
            var blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find end sequence!");
            tokens.RemoveAt(0); // Remove "LOOP"

            // Remove post-condition content (if any)
            if (!hasPreCondition)
            {
                // Note that it's valid in VBScript to have neither a pre- or post-condition and for the loop to continue until an EXIT DO
                // statement is encountered (in which case we will pass a null conditionStatement to the DoBlock). It's not valid for it
                // to have both a pre- and post-condition, so if a pre-condition has already been extracted, don't try to extract a
                // post-condition.
                var postConditionStartToken = base.getToken(tokens,  0, new List<Type> { typeof(AtomToken), typeof(AbstractEndOfStatementToken) });
                if (postConditionStartToken is AtomToken)
                {
                    var doWhile = (postConditionStartToken.Content.ToUpper() == "WHILE");
                    doUntil = (postConditionStartToken.Content.ToUpper() == "UNTIL");
                    if (doWhile || doUntil)
                    {
                        tokens.RemoveAt(0); // Remove WHILE / UNTIL
                        conditionStatement = ExtractConditionFromTokens(tokens);
                    }
                }
            }
            
            // Whether a post-condition has been processed or the construct terminated at the "LOOP" keyword, the next token (if any)
            // must be an end-of-statement
            if (tokens.Count > 0)
            {
                if (tokens[0] is AbstractEndOfStatementToken)
                    tokens.RemoveAt(0);
                else
                    throw new Exception("EndOfStatementToken missing after LOOP");
            }

            return new DoBlock(conditionStatement, hasPreCondition, doUntil, blockContent);
        }

        private Expression ExtractConditionFromTokens(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                throw new ArgumentException("No tokens to extract content from");

            // Add AtomTokens to list until find EndOfStatement (so long as there are any tokens to consume)
            var tokensInCondition = new List<IToken>();
            while (tokens.Count > 0)
            {
                // Once the end-of-statement has been identified, don't try to remove it - leave that up to the caller
                if (base.isEndOfStatement(tokens, 0))
                    break;
                var tokenCondition = base.getToken_AtomOrStringOnly(tokens, 0);
                tokensInCondition.Add(tokenCondition);
                tokens.RemoveAt(0);
            }
            return new Expression(tokensInCondition);
        }
    }
}
