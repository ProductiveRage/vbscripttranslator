using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This token represents a section of script content (not string or comment) that
    /// has not been broken down into its constituent parts yet
    /// </summary>
    [Serializable]
    public class UnprocessedContentToken : IToken
    {
        private string content;
        public UnprocessedContentToken(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            this.content = content;
        }

        public string Content
        {
            get { return this.content; }
        }
    }
}
