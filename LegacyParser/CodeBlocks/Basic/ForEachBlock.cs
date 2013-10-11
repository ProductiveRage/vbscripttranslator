using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ForEachBlock : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private string loopVar;
        private Expression loopSrc;
        private List<ICodeBlock> statements;
        
        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the
        /// doUntil value is not of any consequence
        /// </summary>
        public ForEachBlock(string loopVar, Expression loopSrc, List<ICodeBlock> statements)
        {
            if ((loopVar ?? "").Trim() == "")
                throw new ArgumentException("loopVar is null or  blank");
            if (loopSrc == null)
                throw new ArgumentNullException("loopSrc");
            if (statements == null)
                throw new ArgumentNullException("statements");
            this.loopVar = loopVar;
            this.loopSrc = loopSrc;
            this.statements = statements;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public string LoopVar
        {
            get { return this.loopVar; }
        }

        public Expression LoopSrc
        {
            get { return this.loopSrc; }
        }

        public List<ICodeBlock> Statements
        {
            get { return this.statements; }
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

            // Open statement
            output.Append(indenter.Indent);
            output.Append("For Each ");
            output.Append(this.loopVar);
            output.Append(" In ");
            output.AppendLine(this.loopSrc.GenerateBaseSource(new NullIndenter()));

            // Render inner content
            foreach (ICodeBlock statement in this.statements)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            // Close statement
            output.Append(indenter.Indent + "Next");
            return output.ToString();
        }
    }
}
