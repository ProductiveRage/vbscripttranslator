using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.Tokens
{
    [Serializable]
    public class NumericValueToken : AtomToken
    {
        public NumericValueToken(double value) : base(value.ToString(), WhiteSpaceBehaviourOptions.Disallow) { }
    }
}
