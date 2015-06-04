using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This is a specialisation of the NameToken, where the value "Me" is used it has special meaning
    /// </summary>
    [Serializable]
    public class TargetCurrentClassToken : NameToken
    {
        public TargetCurrentClassToken(int lineIndex) : base("Me", lineIndex) { }
    }
}
