using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class Expression : Statement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// An expression is code that evalutes to a value
        /// </summary>
        public Expression(IEnumerable<IToken> tokens) : base(tokens, CallPrefixOptions.Absent) { }
    }
}
