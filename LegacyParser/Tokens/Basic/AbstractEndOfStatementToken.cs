using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public abstract class AbstractEndOfStatementToken : IToken
    {
        public abstract string Content { get; }
    }
}
