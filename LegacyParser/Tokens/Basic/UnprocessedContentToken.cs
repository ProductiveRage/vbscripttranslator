using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This token represents a section of script content (not string or comment) that
    /// has not been broken down into its constituent parts yet
    /// </summary>
    [Serializable]
    public class UnprocessedContentToken : IToken
    {
        public UnprocessedContentToken(string content, int lineIndex)
        {
            if (string.IsNullOrEmpty(content))
                throw new ArgumentException("Null/blank content specified");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            LineIndex = lineIndex;
            Content = content;
        }

        public string Content { get; private set; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int LineIndex { get; private set; }
    }
}
