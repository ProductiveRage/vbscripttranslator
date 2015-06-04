using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class RandomizeStatement : IHaveNonNestedExpressions
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

        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> IHaveNonNestedExpressions.NonNestedExpressions
        {
            get
            {
                yield return Seed;
            }
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
                output.Append(this.seed.GenerateBaseSource(NullIndenter.Instance));
            }
            return output.ToString();
        }
    }
}
