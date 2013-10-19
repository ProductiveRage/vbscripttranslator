using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class NumericValueToken : AtomToken
    {
        public NumericValueToken(double value) : base(value.ToString(), WhiteSpaceBehaviourOptions.Disallow) { }
    }
}
