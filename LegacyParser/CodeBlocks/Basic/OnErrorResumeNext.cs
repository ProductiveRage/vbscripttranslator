using System;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class OnErrorResumeNext : ICodeBlock
    {
        public OnErrorResumeNext(int lineIndex)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex");

            LineIndex = lineIndex;
        }

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
            return indenter.Indent + "ON ERROR RESUME NEXT";
        }
    }
}
