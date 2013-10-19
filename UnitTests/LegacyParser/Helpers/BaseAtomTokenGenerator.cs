using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.UnitTests.LegacyParser.Helpers
{
    public class BaseAtomTokenGenerator
    {
        public static AtomToken Get(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            var token = AtomToken.GetNewToken(content);
            if (token.GetType() != typeof(AtomToken))
                throw new ArgumentException("Specified content was not mapped to an AtomToken, it was mapped to " + token.GetType());
            return (AtomToken)token;
        }
    }
}
