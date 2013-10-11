using System;
namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class KeyWordToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public KeyWordToken(string content) : base(content)
        {
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (content == null)
                throw new ArgumentNullException("content");
            if (content == "")
                throw new ArgumentException("Blank content specified for KeyWordToken - invalid");
            if (!AtomToken.isMustHandleKeyWord(content) && !AtomToken.isMiscKeyWord(content))
                throw new ArgumentException("Invalid content specified - not a Key Word");
            this.content = content;
        }
    }
}
