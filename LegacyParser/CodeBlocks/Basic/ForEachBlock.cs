using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ForEachBlock : ILoopOverNestedContent, ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private NameToken loopVar;
        private Expression loopSrc;
        private List<ICodeBlock> statements;
        
        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the
        /// doUntil value is not of any consequence
        /// </summary>
        public ForEachBlock(NameToken loopVar, Expression loopSrc, List<ICodeBlock> statements)
        {
            if (loopVar == null)
                throw new ArgumentNullException("loopVar");
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
        public NameToken LoopVar
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

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get { return new ICodeBlock[] { new Expression(new[] { LoopVar }), LoopSrc }.Concat(Statements); }
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
            output.Append(this.loopVar.Content);
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
