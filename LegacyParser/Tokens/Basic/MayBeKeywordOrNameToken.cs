using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// There are tokens that may be reference names or keywords, depending upon context - it is not possible to tell just from their content. Values
    /// (such as "step") are represented by these tokens.
    /// </summary>
    [Serializable]
    public class MayBeKeywordOrNameToken : NameToken
    {
        public MayBeKeywordOrNameToken(string content, int lineIndex) : base(content, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            if (!AtomToken.isContextDependentKeyword(content))
                throw new ArgumentException("Invalid content for a MayBeKeywordOrNameToken");
        }
    }
}
