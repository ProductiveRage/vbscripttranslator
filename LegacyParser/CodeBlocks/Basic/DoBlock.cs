using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class DoBlock : IHaveNestedContent, ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private Expression conditionStatement;
        private bool doUntil;
        private List<ICodeBlock> statements;
        
        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the
        /// doUntil value is not of any consequence
        /// </summary>
        public DoBlock(Expression conditionStatement, bool doUntil, List<ICodeBlock> statements)
        {
            if (statements == null)
                throw new ArgumentNullException("statements");
            this.conditionStatement = conditionStatement;
            this.doUntil = doUntil;
            this.statements = statements;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public Expression Condition
        {
            get { return this.conditionStatement; }
        }

        public bool DoWhileCondition
        {
            get { return !this.doUntil; }
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
            get { return new ICodeBlock[] { Condition }.Concat(Statements); }
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
            output.Append("Do ");
            if (this.doUntil)
                output.Append("Until ");
            else
                output.Append("While ");
            output.AppendLine(this.conditionStatement.GenerateBaseSource(new NullIndenter()));

            // Render inner content
            foreach (ICodeBlock statement in this.statements)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            // Close statement
            output.Append(indenter.Indent + "Loop");
            return output.ToString();
        }
    }
}
