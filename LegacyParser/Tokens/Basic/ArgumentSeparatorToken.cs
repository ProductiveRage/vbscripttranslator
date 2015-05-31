using System;
namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class ArgumentSeparatorToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public ArgumentSeparatorToken(int lineIndex) : base(",", WhiteSpaceBehaviourOptions.Disallow, lineIndex) { }
        }
    }
