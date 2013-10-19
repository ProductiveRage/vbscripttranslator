using System;
namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class OpenBrace : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public OpenBrace(string content) : base(content, WhiteSpaceBehaviourOptions.Disallow)
        {
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (content == null)
                throw new ArgumentNullException("content");
            if (content == "")
                throw new ArgumentException("Blank content specified for OpenBrace - invalid");
            if (!AtomToken.isOpenBrace(content))
                throw new ArgumentException("Invalid content specified - not an Open Brace");
            this.content = content;
        }
    }
}
