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
        private ExitableStatementType statementType;
        public ExitStatement(ExitableStatementType statementType)
        {
            bool isValid = false;
            foreach (object value in Enum.GetValues(typeof(ExitableStatementType)))
            {
                if (value.Equals(statementType))
                {
                    isValid = true;
                    break;
                }
            }
            if (!isValid)
				throw new ArgumentException("Invalid statementType value specified [" + statementType.ToString() + "]");
            this.statementType = statementType;
        }

        public enum ExitableStatementType
        {
			Do,
			For,
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
            return indenter.Indent + "Exit " + this.statementType.ToString();
        }
    }
}
