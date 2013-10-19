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
        public CloseBrace(string content) : base(content, WhiteSpaceBehaviourOptions.Disallow)
        {
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (content == null)
                throw new ArgumentNullException("content");
            if (content == "")
                throw new ArgumentException("Blank content specified for CloseBrace - invalid");
            if (!AtomToken.isCloseBrace(content))
                throw new ArgumentException("Invalid content specified - not a Close Brace");
            this.content = content;
        }
    }
}
