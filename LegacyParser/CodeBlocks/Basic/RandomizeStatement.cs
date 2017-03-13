using System;
using System.Collections.Generic;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class RandomizeStatement : IHaveNonNestedExpressions
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
		public RandomizeStatement(int lineIndex, Expression seedIfAny)
        {
			if (lineIndex < 0)
				throw new ArgumentOutOfRangeException(nameof(lineIndex));

			LineIndex = LineIndex;
			SeedIfAny = seedIfAny;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
		public int LineIndex { get; }

        /// <summary>
        /// Note: This may be null
        /// </summary>
		public Expression SeedIfAny { get; }

        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> IHaveNonNestedExpressions.NonNestedExpressions
        {
            get
            {
				if (SeedIfAny != null)
					yield return SeedIfAny;
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
			if (SeedIfAny != null)
            {
                output.Append(" ");
				output.Append(SeedIfAny.GenerateBaseSource(NullIndenter.Instance));
            }
            return output.ToString();
        }
    }
}
