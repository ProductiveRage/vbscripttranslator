using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class EndOfStatementSameLineToken : AbstractEndOfStatementToken
    {
        public EndOfStatementSameLineToken(int lineIndex) : base(lineIndex) { }

        public override string Content
        {
            get { return ""; }
        }
    }
}
