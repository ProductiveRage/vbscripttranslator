using System;
using CSharpSupport;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a single date literal section
    /// </summary>
    [Serializable]
    public class DateLiteralToken : IToken
    {
        public DateLiteralToken(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            // This is done just to ensure that the content is not nonsense. It will be interpreted at runtime so that if the translated code is run
            // with a different culture to the translation process then that culture is respected when the date is parsed.
            try
            {
                DateParser.Default.Parse(content);
            }
            catch (Exception e)
            {
                throw new ArgumentException("Invalid date literal content", e);
            }

            Content = content;
            LineIndex = lineIndex;
        }

        /// <summary>
        /// This will not include the hashes in the value
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
