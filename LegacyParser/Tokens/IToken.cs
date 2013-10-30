namespace VBScriptTranslator.LegacyParser.Tokens
{
    public interface IToken
    {
        string Content { get; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        int LineIndex { get; }
    }
}
