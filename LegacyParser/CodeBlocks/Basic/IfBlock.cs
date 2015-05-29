using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class IfBlock : IHaveNestedContent
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public IfBlock(IEnumerable<IfBlockSegment> clauses)
        {
            if (clauses == null)
                throw new ArgumentNullException("clauses");

            var clausesArray = clauses.ToArray();
            if (!clausesArray.Any())
                throw new ArgumentException("Empty clauses set specified - invalid");
            if (clausesArray.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in clauses set");
            
            var numberOfElseSegments = clausesArray.Count(c => c is IfBlockElseSegment);
            if (numberOfElseSegments > 1)
                throw new ArgumentException("There may never be more than one IfBlockElseSegment");
            if (numberOfElseSegments == 1)
            {
                if ((clausesArray.Length == 1) || !(clausesArray.Last() is IfBlockElseSegment))
                    throw new ArgumentException("If an IfBlockElseSegment is present, it must be the last clause (and is not allowed if there is only a single clause");
            }

            var firstInvalidSegmentIfAny = clausesArray.FirstOrDefault(c => !(c is IfBlockConditionSegment) && !(c is IfBlockElseSegment));
            if (firstInvalidSegmentIfAny != null)
                throw new ArgumentException("Unsupported segment type: " + firstInvalidSegmentIfAny.GetType());

            ConditionalClauses = clausesArray.OfType<IfBlockConditionSegment>();
            OptionalElseClause = (IfBlockElseSegment)clausesArray.FirstOrDefault(c => c is IfBlockElseSegment);
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null, empty or contain any nulls
        /// </summary>
        public IEnumerable<IfBlockConditionSegment> ConditionalClauses { get; private set; }
        
        /// <summary>
        /// This will be null if there was no fallback clause
        /// </summary>
        public IfBlockElseSegment OptionalElseClause { get; private set; }

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get
            {
                foreach (var conditionalClause in ConditionalClauses)
                {
                    yield return conditionalClause.Condition;
                    foreach (var statement in conditionalClause.Statements)
                        yield return statement;
                }
                if (OptionalElseClause == null)
                    yield break;
                foreach (var statement in OptionalElseClause.Statements)
                    yield return statement;
            }
        }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        public interface IfBlockSegment
        {
            IEnumerable<ICodeBlock> Statements { get; }
        }

        public class IfBlockConditionSegment : IfBlockSegment
        {
            public IfBlockConditionSegment(Expression conditionStatement, IEnumerable<ICodeBlock> statements)
            {
                if (conditionStatement == null)
                    throw new ArgumentNullException("conditionStatement");
                if (statements == null)
                    throw new ArgumentNullException("statements");

                Statements = statements.ToList().AsReadOnly();
                if (Statements.Any(s => s == null))
                    throw new ArgumentException("Null reference encountered in statements set");
                Condition = conditionStatement;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public Expression Condition { get; private set; }

            /// <summary>
            /// This will never be null or contain any nulls
            /// </summary>
            public IEnumerable<ICodeBlock> Statements { get; private set; }
        }
        
        public class IfBlockElseSegment : IfBlockSegment
        {
            public IfBlockElseSegment(IEnumerable<ICodeBlock> statements)
            {
                if (statements == null)
                    throw new ArgumentNullException("statements");
                Statements = statements.ToList().AsReadOnly();
                if (Statements.Any(s => s == null))
                    throw new ArgumentException("Null reference encountered in statements set");
            }

            /// <summary>
            /// This will never be null or contain any nulls
            /// </summary>
            public IEnumerable<ICodeBlock> Statements { get; private set; }
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
            var output = new StringBuilder();

            var allClauses = ConditionalClauses.Cast<IfBlockSegment>().ToList();
            if (OptionalElseClause != null)
                allClauses.Add(OptionalElseClause);

            for (int index = 0; index < allClauses.Count; index++)
            {
                // Render branch start: IF / ELSEIF / ELSE
                var segment = allClauses[index];
                if (segment is IfBlockConditionSegment)
                {
                    output.Append(indenter.Indent);
                    if (index == 0)
                        output.Append("IF ");
                    else
                        output.Append("ELSEIF ");
                    output.Append(
                        ((IfBlockConditionSegment)segment).Condition.GenerateBaseSource(new NullIndenter())
                    );
                    output.AppendLine(" THEN");
                }
                else
                    output.AppendLine(indenter.Indent + "ELSE");

                // Render branch content
                foreach (var statement in segment.Statements)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));
            }
            output.Append(indenter.Indent + "END IF");
            return output.ToString();
        }
    }
}
