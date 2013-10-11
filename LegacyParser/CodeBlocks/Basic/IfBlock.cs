using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class IfBlock : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private List<IfBlockSegment> content;
        public IfBlock(List<IfBlockSegment> content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            foreach (IfBlockSegment contentInner in content)
            {
                if (contentInner == null)
                    throw new ArgumentException("Encountered null inner content");
            }
            this.content = content;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public List<IfBlockSegment> Content
        {
            get { return this.content; }
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
            public Expression Condition { get; private set; }
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
            StringBuilder output = new StringBuilder();
            for (int index = 0; index < this.content.Count; index++)
            {
                // Render branch start: IF / ELSEIF / ELSE
                IfBlockSegment segment = this.content[index];
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
                foreach (ICodeBlock statement in segment.Statements)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));
            }
            output.Append(indenter.Indent + "END IF");
            return output.ToString();
        }
    }
}
