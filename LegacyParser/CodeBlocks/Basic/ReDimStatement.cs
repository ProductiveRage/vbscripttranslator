using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// Note: If VBScript encounters one of these without a previous DIM for the same variable then it will not raise an error, even if OPTION
    /// EXPLICIT was specified - this may be considered to declare a variable initially, it does not necessarily have to re-declare it.
    /// </summary>
    [Serializable]
    public class ReDimStatement : BaseDimStatement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public ReDimStatement(bool preseve, IEnumerable<DimVariable> variables) : base(variables)
        {
            // Ensure that all variables have at least one dimension (VBScript code will not compile if this is not the case and the assumption
            // is that we're dealing with valid code - it would be a compile time error, not a runtime error that could be masked with On Error
            // Resume Next)
            if (Variables.Any(v => (v == null) || !v.Dimensions.Any()))
                throw new ArgumentException("There must be at least one argument for all variables specified in a ReDim statement");

            Preserve = preseve;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // - Note: Variables is already exposed through DimStatement base class
        // =======================================================================================
        public bool Preserve { get; private set; }

        /// <summary>
        /// This will never be null nor contain any nulls (though it may be an empty set), any variables will have at least one dimension
        /// (since VBScript code will not compile statements such as "ReDim a", it must be something like "ReDim a(0)")
        /// </summary>
        public new IEnumerable<DimVariable> Variables { get { return base.Variables; } }

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
            if (Preserve)
                output.Append("Preserve ");
            output.Append(baseContent.Substring(4));
            return output.ToString();
        }
    }
}
