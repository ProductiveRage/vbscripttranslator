using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class CommentToken : IToken
    {
        private string content;
        public CommentToken(string content)
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
