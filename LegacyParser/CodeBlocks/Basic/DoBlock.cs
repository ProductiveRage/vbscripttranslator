using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class DoBlock : ILoopOverNestedContent, ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the isPreCondition and doUntil constructor arguments are of no consequence
        /// </summary>
        public DoBlock(Expression conditionIfAny, bool isPreCondition, bool doUntil, IEnumerable<ICodeBlock> statements)
        {
            if (statements == null)
                throw new ArgumentNullException("statements");

            ConditionIfAny = conditionIfAny;
            IsPreCondition = isPreCondition;
            IsDoWhileCondition = !doUntil;
            Statements = statements.ToList().AsReadOnly();
            if (Statements.Any(s => s == null))
                throw new ArgumentException("Null reference encountered in statements set");
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This may be null since VBScript supports DO..WHILE loops with no constraint
        /// </summary>
        public Expression ConditionIfAny { get; private set; }

        public bool IsPreCondition { get; private set; }

        /// <summary>
        /// If this is true then for construct is of the for DO WHILE..LOOP or DO..LOOP WHILE, as opposed to DO UNTIL..LOOP or DO..LOOP UNTIL
        /// </summary>
        public bool IsDoWhileCondition { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set
        /// </summary>
        public IEnumerable<ICodeBlock> Statements { get; private set; }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get
            {
                var statementBlocks = Statements.ToList();
                if (ConditionIfAny != null)
                {
                    if (IsPreCondition)
                        statementBlocks.Insert(0, ConditionIfAny);
                    else
                        statementBlocks.Add(ConditionIfAny);
                }
                return statementBlocks;
            }
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            var output = new StringBuilder();

            // Open statement (with condition if this construct has a pre condition)
            output.Append("Do");
            if (IsPreCondition && (ConditionIfAny != null))
            {
                output.Append(" ");
                if (IsDoWhileCondition)
                    output.Append("While ");
                else
                    output.Append("Until ");
                output.AppendLine(ConditionIfAny.GenerateBaseSource(new NullIndenter()));
            }
            else
                output.AppendLine();

            // Render inner content
            foreach (var statement in Statements)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            // Close statement (with condition if this construct has a pre condition)
            output.Append(indenter.Indent + "Loop");
            if (!IsPreCondition && (ConditionIfAny != null))
            {
                output.Append(" ");
                if (IsDoWhileCondition)
                    output.Append("While ");
                else
                    output.Append("Until ");
                output.Append(ConditionIfAny.GenerateBaseSource(new NullIndenter()));
            }
            return output.ToString();
        }
    }
}
