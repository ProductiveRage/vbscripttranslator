using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ForBlock : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private string loopVar;
        private Expression loopFrom;
        private Expression loopTo;
        private Expression loopStep;
        private List<ICodeBlock> statements;

        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the
        /// doUntil value is not of any consequence
        /// </summary>
        public ForBlock(string loopVar, Expression loopFrom, Expression loopTo, Expression loopStep, List<ICodeBlock> statements)
        {
            if ((loopVar ?? "").Trim() == "")
                throw new ArgumentException("loopVar is null or  blank");
            if (loopFrom == null)
                throw new ArgumentNullException("loopFrom");
            if (loopTo == null)
                throw new ArgumentNullException("loopTo");
            if (statements == null)
                throw new ArgumentNullException("statements");
            this.loopVar = loopVar;
            this.loopFrom = loopFrom;
            this.loopTo = loopTo;
            this.loopStep = loopStep;
            this.statements = statements;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public string LoopVar
        {
            get { return this.loopVar; }
        }

        public Expression LoopFrom
        {
            get { return this.loopFrom; }
        }

        public Expression LoopTo
        {
            get { return this.loopTo; }
        }

        /// <summary>
        /// Note: This may be null
        /// </summary>
        public Expression LoopStep
        {
            get { return this.loopStep; }
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
            output.Append("For ");
            output.Append(this.loopVar);
            output.Append(" = ");
            output.Append(this.loopFrom.GenerateBaseSource(new NullIndenter()));
            output.Append(" To ");
            output.Append(this.loopTo.GenerateBaseSource(new NullIndenter()));
            if (this.loopStep != null)
            {
                output.Append(" Step ");
                output.Append(this.LoopStep.GenerateBaseSource(new NullIndenter()));
            }
            output.AppendLine("");

            // Render inner content
            foreach (ICodeBlock statement in this.statements)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            // Close statement
            output.Append(indenter.Indent + "Next");
            return output.ToString();
        }
    }
}
