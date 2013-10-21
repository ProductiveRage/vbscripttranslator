using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class IfBlock : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public IfBlock(IEnumerable<IfBlockSegment> clauses)
        {
            if (clauses == null)
                throw new ArgumentNullException("clauses");

            Clauses = clauses.ToList().AsReadOnly();
            if (!Clauses.Any())
                throw new ArgumentException("Empty clauses set specified - invalid");
            if (Clauses.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in clauses set");
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null, empty or contain any nulls
        /// </summary>
        public IEnumerable<IfBlockSegment> Clauses { get; private set; }

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
            var numberOfClauses = Clauses.Count();
            for (int index = 0; index < numberOfClauses; index++)
            {
                // Render branch start: IF / ELSEIF / ELSE
                var segment = Clauses.ElementAt(index);
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
