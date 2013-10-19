using System;
using System.Collections.Generic;
using System.Text;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ClassBlock : ICodeBlock, IDefineScope
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private string className;
        private List<ICodeBlock> statements;
        public ClassBlock(string className, List<ICodeBlock> statements)
        {
            if ((className ?? "").Trim() == "")
                throw new ArgumentException("className is null or blank");
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
            return base.ToString() + ":" + this.className;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public string Name
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
            output.AppendLine(indenter.Indent + "Class " + this.className);
            foreach (ICodeBlock block in this.statements)
                output.AppendLine(block.GenerateBaseSource(indenter.Increase()));
            output.Append(indenter.Indent + "End Class");
            return output.ToString();
        }
    }
}
