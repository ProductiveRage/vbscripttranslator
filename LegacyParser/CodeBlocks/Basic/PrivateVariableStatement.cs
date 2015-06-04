using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class PrivateVariableStatement : DimStatement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public PrivateVariableStatement(List<DimVariable> variables) : base(variables) { }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public override string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            // Grab content from DimStatement..
            string baseContent = base.GenerateBaseSource(NullIndenter.Instance);
            if ((baseContent == null)
            || (baseContent.Length < 4)
            || (baseContent.Substring(0, 4).ToUpper() != "DIM "))
                throw new Exception("Unexpected content from base class");

            // .. and change to be ReDim (add in Preserve keyword, if required)
            StringBuilder output = new StringBuilder();
            output.Append(indenter.Indent);
            output.Append("Private ");
            output.Append(baseContent.Substring(4));
            return output.ToString();
        }
    }
}
