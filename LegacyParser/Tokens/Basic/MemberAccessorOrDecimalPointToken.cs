using System;
namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a member accessor, such as a function or property - eg. the period in target.Name or target.Method() - or the
    /// decimal point in a number - eg. in "1.1"
    /// </summary>
    [Serializable]
    public class MemberAccessorOrDecimalPointToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public MemberAccessorOrDecimalPointToken(string content) : base(content, WhiteSpaceBehaviourOptions.Disallow)
        {
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (content == null)
                throw new ArgumentNullException("content");
            if (content == "")
                throw new ArgumentException("Blank content specified for MemberAccessorToken - invalid");
            if (!AtomToken.isMemberAccessor(content))
                throw new ArgumentException("Invalid content specified - not a MemberAccessor");
            this.content = content;
        }
    }
}
