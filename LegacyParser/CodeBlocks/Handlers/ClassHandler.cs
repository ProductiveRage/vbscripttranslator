using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class ClassHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count < 3)
                return null;

            // Look for start of function declaration
            string[] matchPattern = new string[] { "CLASS" };
            if (!base.checkAtomTokenPattern(tokens, matchPattern, false))
                return null;
            if (!(tokens[1] is AtomToken))
                return null;
            if (!(tokens[2] is AbstractEndOfStatementToken))
                return null;
            string className = tokens[1].Content;
            tokens.RemoveRange(0, 3);

            // Get function content
            string[] endSequenceMet;
            var endSequences = new List<string[]>()
            {
                new string[] { "END", "CLASS" }
            };
            var codeBlockHandler = new CodeBlockHandler(endSequences);
            var functionContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find encounter end sequence!");
            
            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if ((tokens.Count > 0) && (!(tokens[0] is AbstractEndOfStatementToken)))
                throw new Exception("EndOfStatementToken missing after END CLASS");
            else
                tokens.RemoveAt(0);

            // Return Function code block instance
            return new ClassBlock(className, functionContent);
        }
    }
}
