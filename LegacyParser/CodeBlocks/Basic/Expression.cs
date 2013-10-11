using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class Expression : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// An expression is code that evalutes to a value
        /// </summary>
        public Expression(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            Tokens = tokens.ToList().AsReadOnly();
            foreach (var token in Tokens)
            {
                if (token == null)
                    throw new ArgumentException("Null token passed into Expression constructor");
                if ((!(token is AtomToken)) && (!(token is StringToken)))
                    throw new ArgumentException("Expression may only be initialised with Atom and String tokens");
            }
            if (!Tokens.Any())
                throw new ArgumentException("Expression initialised with empty token stream - invalid");
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null, empty nor contain any nulls
        /// </summary>
        public IEnumerable<IToken> Tokens { get; private set; }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            StringBuilder output = new StringBuilder();
            output.Append(indenter.Indent);
            var index = 0;
            var tokenCount = Tokens.Count();
            foreach (var token in Tokens)
            {
                if (token is StringToken)
                    output.Append("\"" + token.Content + "\"");
                else
                    output.Append(token.Content);
                if (index < (tokenCount - 1))
                    output.Append(" ");
                index++;
            }
            return output.ToString();
        }
    }
}
