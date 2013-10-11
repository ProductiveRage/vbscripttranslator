using System;
namespace VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering
{
    public class SourceIndentHandler : ISourceIndentHandler
    {
        private int depth;
        public SourceIndentHandler() : this(0) {}
        private SourceIndentHandler(int depth)
        {
            if (depth < 0)
                throw new ArgumentException("Negative depth - invalid");
            this.depth = depth;
        }

        public ISourceIndentHandler Increase()
        {
            return new SourceIndentHandler(this.depth + 1);
        }

        public ISourceIndentHandler Decrease()
        {
            return new SourceIndentHandler(this.depth - 1);
        }

        public string Indent
        {
            get
            {
                return new String(' ', this.depth * 2);
            }
        }
    }
}
