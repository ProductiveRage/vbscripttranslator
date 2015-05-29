using System;
using CSharpSupport;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a single date literal section. It can not be known at this time whether the value is a valid date or not, it may vary depending upon
    /// culture (eg. "1 Jan 2015" is valid when the English language is used but not for other languages). Before performing any work, the translated program
    /// must verify that all date literals are valid - this is equivalent to the VBScript interpreter validating dates when the script is first read and raising
    /// a syntax error if any date literal is invalid.
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

            Content = content;
            LineIndex = lineIndex;
        }

        /// <summary>
        /// This will not include the hashes in the value (it is not possible to be sure that this is valid content, since it may vary by culture - which may be
        /// different when the translated program is run than it was during the translation process)
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
