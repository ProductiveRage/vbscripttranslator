namespace VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering
{
    public interface ISourceIndentHandler
    {
        ISourceIndentHandler Increase();
        ISourceIndentHandler Decrease();
        string Indent { get; }
    }
}
