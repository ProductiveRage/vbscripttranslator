using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a single string section
    /// </summary>
    [Serializable]
    public class StringToken : IToken
    {
        public StringToken(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            Content = content;
            LineIndex = lineIndex;
        }

        /// <summary>
        /// This will not include the quotes in the value
        /// </summary>
        public string Content { get; private set; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int LineIndex { get; private set; }

        public override string ToString()
        {
            return base.ToString() + ":" + Content;
        }
    }
}
