using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public abstract class AbstractEndOfStatementToken : IToken
    {
        public AbstractEndOfStatementToken(int lineIndex)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            LineIndex = lineIndex;
        }

        public abstract string Content { get; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int LineIndex { get; private set; }
    }
}
