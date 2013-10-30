using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class EndOfStatementNewLineToken : AbstractEndOfStatementToken
    {
        public EndOfStatementNewLineToken(int lineIndex) : base(lineIndex) { }

        public override string Content
        {
            get { return ""; }
        }
    }
}
