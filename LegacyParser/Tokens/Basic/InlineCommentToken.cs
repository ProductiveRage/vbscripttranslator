using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class InlineCommentToken : CommentToken
    {
        public InlineCommentToken(string content) : base(content) { }
    }
}
