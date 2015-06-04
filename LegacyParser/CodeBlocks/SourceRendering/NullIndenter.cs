namespace VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering
{
    public class NullIndenter : ISourceIndentHandler
    {
        public static NullIndenter _instance = new NullIndenter();
        public static NullIndenter Instance { get { return _instance; } }
        private NullIndenter() { }

        public ISourceIndentHandler Increase()
        {
            return NullIndenter.Instance;
        }

        public ISourceIndentHandler Decrease()
        {
            return NullIndenter.Instance;
        }

        public string Indent
        {
            get { return ""; }
        }
    }
}
