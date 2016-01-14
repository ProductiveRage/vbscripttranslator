using System;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class InlineCommentStatement : CommentStatement
    {
        public InlineCommentStatement(string content, int lineIndex) : base(content, lineIndex) { }
    }
}
