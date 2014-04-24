using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class NumericValueToken : AtomToken
    {
        public NumericValueToken(double value, int lineIndex) : base(value.ToString(), WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            Value = value;
        }

        public double Value { get; private set; }
    }
}
