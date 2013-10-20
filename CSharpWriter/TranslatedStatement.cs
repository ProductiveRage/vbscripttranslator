using System;

namespace CSharpWriter
{
    public class TranslatedStatement
    {
        public TranslatedStatement(string content, int indentationDepth)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (content != content.Trim())
                throw new ArgumentException("content may be blank but may not have any leading or trailing whitespace");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            Content = content;
            IndentationDepth = indentationDepth;
        }

        /// <summary>
        /// This will never be null, though it may be blank if it represents a blank line. It will never have any leading or trailing whitespace.
        /// </summary>
        public string Content { get; private set; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int IndentationDepth { get; private set; }
    }
}
