using System;
namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class CloseBrace : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public CloseBrace() : base(")", WhiteSpaceBehaviourOptions.Disallow) { }
    }
}
