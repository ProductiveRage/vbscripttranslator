using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ForBlock : ILoopOverNestedContent, ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private NameToken loopVar;
        private Expression loopFrom;
        private Expression loopTo;
        private Expression loopStep;
        private List<ICodeBlock> statements;

        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the
        /// doUntil value is not of any consequence
        /// </summary>
        public ForBlock(NameToken loopVar, Expression loopFrom, Expression loopTo, Expression loopStep, List<ICodeBlock> statements)
        {
            if (loopVar == null)
                throw new ArgumentNullException("loopVar");
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
        public NameToken LoopVar
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

        public IEnumerable<ICodeBlock> Statements
        {
            get { return this.statements.AsReadOnly(); }
        }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get
            {
                return new ICodeBlock[] { new Expression(new[] { LoopVar }), LoopFrom, LoopTo, LoopStep }
                    .Where(b => b != null) // Ignore a null LoopStep (this is a valid configuration but we can't have nulls in the data returned here)
                    .Concat(Statements);
            }
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
            output.Append(this.loopVar.Content);
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
