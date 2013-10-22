using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ExitStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private ExitableStatementType loopType;
        public ExitStatement(ExitableStatementType loopType)
        {
            bool isValid = false;
            foreach (object value in Enum.GetValues(typeof(ExitableStatementType)))
            {
                if (value.Equals(loopType))
                {
                    isValid = true;
                    break;
                }
            }
            if (!isValid)
                throw new ArgumentException("Invalid type value specified [" + loopType.ToString() + "]");
            this.loopType = loopType;
        }

        public enum ExitableStatementType
        {
            For,
            Do,
			Function,
			Property
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
            return indenter.Indent + "Exit " + this.loopType.ToString();
        }
    }
}
