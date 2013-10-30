using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.Tokens
{
    /// <summary>
    /// This represents a member accessor, such as a function or property - eg. the period in target.Name or target.Method()
    /// This is more specific than MemberAccessorOrDecimalPointToken, which may be an accessor or a decimal place (eg. the "." in "1.1")
    /// </summary>
    [Serializable]
    public class MemberAccessorToken : MemberAccessorOrDecimalPointToken
    {
        public MemberAccessorToken(int lineIndex) : base(".", lineIndex) { }
    }
}
