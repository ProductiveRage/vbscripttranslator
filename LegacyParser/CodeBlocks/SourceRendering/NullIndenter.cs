namespace VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering
{
    public class NullIndenter : ISourceIndentHandler
    {
        public ISourceIndentHandler Increase()
        {
            return new NullIndenter();
        }

        public ISourceIndentHandler Decrease()
        {
            return new NullIndenter();
        }

        public string Indent
        {
            get { return ""; }
        }
    }
}
