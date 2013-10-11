using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class RandomizeStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private Expression seed;
        public RandomizeStatement(Expression seed)
        {
            this.seed = seed;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// Note: This may be null
        /// </summary>
        public Expression Seed
        {
            get { return this.seed; }
        }

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
            output.Append("Randomize");
            if (this.seed != null)
            {
                output.Append(" ");
                output.Append(this.seed.GenerateBaseSource(new NullIndenter()));
            }
            return output.ToString();
        }
    }
}
