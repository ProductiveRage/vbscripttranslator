using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.UnitTests.Shared
{
    public static class Misc
    {
        public static AtomToken GetAtomToken(string content)
        {
            var token = AtomToken.GetNewToken(content);
            if (token.GetType() != typeof(AtomToken))
                throw new ArgumentException("Specified content was not mapped to an AtomToken, it was mapped to " + token.GetType());
            return (AtomToken)token;
        }
    }
}
