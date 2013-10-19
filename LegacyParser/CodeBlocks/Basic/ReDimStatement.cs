﻿using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// Note: If VBScript encounters one of these without a previous DIM for the same variable then it will not raise an error, even if OPTION
    /// EXPLICIT was specified - this may be considered to declare a variable initially, it does not necessarily have to re-declare it.
    /// </summary>
    [Serializable]
    public class ReDimStatement : DimStatement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private bool preseve;
        public ReDimStatement(bool preseve, List<DimVariable> variables) : base(variables)
        {
            this.preseve = preseve;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // - Note: Variables is already exposed through DimStatement base class
        // =======================================================================================
        public bool Preserve
        {
            get { return this.preseve; }
        }

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
            string baseContent = base.GenerateBaseSource(new NullIndenter());
            if ((baseContent == null)
            || (baseContent.Length < 4)
            || (baseContent.Substring(0, 4).ToUpper() != "DIM "))
                throw new Exception("Unexpected content from base class");

            // .. and change to be ReDim (add in Preserve keyword, if required)
            StringBuilder output = new StringBuilder();
            output.Append("ReDim ");
            if (this.preseve)
                output.Append("Preserve ");
            output.Append(baseContent.Substring(4));
            return output.ToString();
        }
    }
}
