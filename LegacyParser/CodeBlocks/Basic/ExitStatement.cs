using System;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ExitStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public ExitStatement(ExitableStatementType statementType, int lineIndex)
		{
            if (!Enum.IsDefined(typeof(ExitableStatementType), statementType))
				throw new ArgumentException("Invalid statementType value specified [" + statementType.ToString() + "]");
			if (lineIndex < 0)
				throw new ArgumentOutOfRangeException("lineIndex");

			StatementType = statementType;
			LineIndex = lineIndex;
		}

		public ExitableStatementType StatementType { get; private set; }

		/// <summary>
		/// This will always be zero or greater
		/// </summary>
		public int LineIndex { get; private set; }

		public enum ExitableStatementType
        {
			Do,
			For,
			Function,
			Property,
            Sub
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
            return indenter.Indent + "Exit " + StatementType.ToString();
        }
    }
}
