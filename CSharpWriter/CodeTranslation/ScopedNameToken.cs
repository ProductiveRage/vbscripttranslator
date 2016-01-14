using System;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation
{
    public class ScopedNameToken : NameToken
    {
        public ScopedNameToken(string content, int lineIndex, ScopeLocationOptions scopeLocation) : base(content, lineIndex)
        {
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");

            ScopeLocation = scopeLocation;
        }

        public ScopeLocationOptions ScopeLocation { get; private set; }
    }
}
