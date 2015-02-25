using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
    public class FuncByRefMapping
    {
        public FuncByRefMapping(NameToken from, CSharpName to)
        {
            if (from == null)
                throw new ArgumentNullException("from");
            if (to == null)
                throw new ArgumentNullException("to");

            From = from;
            To = to;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken From { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public CSharpName To { get; private set; }
    }
}
