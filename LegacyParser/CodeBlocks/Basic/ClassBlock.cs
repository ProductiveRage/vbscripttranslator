using System;
using System.Collections.Generic;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ClassBlock : ICodeBlock, IDefineScope
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private NameToken className;
        private List<ICodeBlock> statements;
        public ClassBlock(NameToken className, List<ICodeBlock> statements)
        {
            if (className == null)
                throw new ArgumentNullException("className");
            if (statements == null)
                throw new ArgumentNullException("statements");

            foreach (ICodeBlock block in statements)
            {
                if (block == null)
                    throw new ArgumentException("Null block in statements");
            }

            this.className = className;
            this.statements = statements;
        }

        public override string ToString()
        {
            return base.ToString() + ":" + this.className.Content;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public NameToken Name
        {
            get { return this.className; }
        }

        public List<ICodeBlock> Statements
        {
            get { return this.statements; }
        }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get { return this.statements.AsReadOnly(); }
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            StringBuilder output = new StringBuilder();
            output.AppendLine(indenter.Indent + "Class " + this.className.Content);
            foreach (ICodeBlock block in this.statements)
                output.AppendLine(block.GenerateBaseSource(indenter.Increase()));
            output.Append(indenter.Indent + "End Class");
            return output.ToString();
        }
    }
}
