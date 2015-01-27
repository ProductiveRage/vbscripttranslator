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
        public BuiltInFunctionToken(string content, int lineIndex) : base(content, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (string.IsNullOrWhiteSpace(content))
                throw new ArgumentException("Null/blank content specified");
            if (!AtomToken.isVBScriptFunction(content))
                throw new ArgumentException("Invalid content specified - not a VBScript function");
        }

        /// <summary>
        /// Is this a function that will always return a numeric value (or raise an error)? This will not return true for functions such as ABS
        /// which return VBScript Null in some cases, it will only apply to functions which always return a "true" number (eg. CDBL).
        /// </summary>
        public bool GuaranteedToReturnNumericContent
        {
            get { return AtomToken.isVBScriptFunctionThatAlwaysReturnsNumericContent(Content); }
        }
    }
}
