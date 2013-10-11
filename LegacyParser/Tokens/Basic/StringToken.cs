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
        private string content;
        public StringToken(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            this.content = content;
        }

        public string Content
        {
            get { return this.content; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + this.content;
        }
    }
}
