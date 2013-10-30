using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class CommentToken : IToken
    {
        public CommentToken(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");

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
