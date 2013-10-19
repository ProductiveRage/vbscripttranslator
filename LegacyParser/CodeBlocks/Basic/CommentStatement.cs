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
        private string content;
        public CommentStatement(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            this.content = content;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public string Content
        {
            get { return this.content; }
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
            if (this.content.Trim() == "")
                return "";
            return indenter.Indent + "'" + this.content;
        }
    }
}
