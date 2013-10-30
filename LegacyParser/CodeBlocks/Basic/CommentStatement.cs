using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class CommentStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public CommentStatement(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (content.Contains("\n"))
                throw new ArgumentException("The content may not include any line returns");

            Content = content.TrimEnd();
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null or contain any line returns. It may be blank and may have leading whitespace (though it won't have
        /// any trailing whitespace).
        /// </summary>
        public string Content { get; private set; }

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
