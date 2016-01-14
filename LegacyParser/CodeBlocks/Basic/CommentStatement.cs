using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class CommentStatement : INonExecutableCodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public CommentStatement(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (content.Contains("\n"))
                throw new ArgumentException("The content may not include any line returns");
			if (lineIndex < 0)
				throw new ArgumentOutOfRangeException("lineIndex");

			Content = content.TrimEnd();
			LineIndex = lineIndex;
		}

		// =======================================================================================
		// PUBLIC DATA ACCESS
		// =======================================================================================
		/// <summary>
		/// This will never be null or contain any line returns. It may be blank and may have leading whitespace (though it won't have
		/// any trailing whitespace).
		/// </summary>
		public string Content { get; private set; }

		/// <summary>
		/// This will always be zero or greater
		/// </summary>
		public int LineIndex { get; private set; }

		// =======================================================================================
		// VBScript BASE SOURCE RE-GENERATION
		// =======================================================================================
		/// <summary>
		/// Re-generate equivalent VBScript source code for this block - there
		/// should not be a line return at the end of the content
		/// </summary>
		public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            if (Content.Trim() == "")
                return "";
            return indenter.Indent + "'" + Content;
        }
    }
}
