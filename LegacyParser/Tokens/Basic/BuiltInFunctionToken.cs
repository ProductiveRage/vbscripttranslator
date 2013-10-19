using System;
namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class BuiltInFunctionToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public BuiltInFunctionToken(string content) : base(content, WhiteSpaceBehaviourOptions.Disallow)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (string.IsNullOrWhiteSpace(content))
                throw new ArgumentException("Null/blank content specified");
            if (!AtomToken.isVBScriptFunction(content))
                throw new ArgumentException("Invalid content specified - not a VBScript function");
        }
    }
}
