using System;
using System.Text;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class SelectBlock : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private Expression expression;
        private List<CommentStatement> openingComments;
        private List<CaseBlockSegment> content;

        public SelectBlock(
            Expression expression,
            List<CommentStatement> openingComments,
            List<CaseBlockSegment> content)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");
            this.expression = expression;
            if ((openingComments == null) || (openingComments.Count == 0))
                this.openingComments = null;
            else
                this.openingComments = openingComments;
            if ((content == null) || (content.Count == 0))
                this.content = null;
            else
            {
                foreach (CaseBlockSegment contentInner in content)
                {
                    if (contentInner == null)
                        throw new ArgumentException("Encountered null inner content");
                }
                this.content = content;
            }
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public Expression Expression
        {
            get { return this.expression; }
        }

        public List<CommentStatement> OpeningComments
        {
            get { return this.openingComments; }
        }

        public List<CaseBlockSegment> Content
        {
            get { return this.content; }
        }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        public interface CaseBlockSegment
        {
            List<ICodeBlock> Statements { get; }
        }

        public class CaseBlockExpressionSegment : CaseBlockSegment
        {
            private List<Expression> values;
            private List<ICodeBlock> statements;
            public CaseBlockExpressionSegment(List<Expression> values, List<ICodeBlock> statements)
            {
                if (values == null)
                    throw new ArgumentNullException("values");
                if (values.Count == 0)
                    throw new ArgumentException("values is an empty list - invalid");
                if (statements == null)
                    throw new ArgumentNullException("statements");
                this.values = values;
                this.statements = statements;
            }
            public List<Expression> Values { get { return this.values; } }
            public List<ICodeBlock> Statements { get { return this.statements; } }
        }
        
        public class CaseBlockElseSegment : CaseBlockSegment
        {
            private List<ICodeBlock> statements;
            public CaseBlockElseSegment(List<ICodeBlock> statements)
            {
                if (statements == null)
                    throw new ArgumentNullException("statements");
                this.statements = statements;
            }
            public List<ICodeBlock> Statements{ get { return this.statements; } }
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
            
            output.Append(indenter.Indent + "SELECT CASE ");
            output.AppendLine(this.expression.GenerateBaseSource(new NullIndenter()));

            if (this.openingComments != null)
            {
                foreach (CommentStatement statement in this.openingComments)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));
                output.AppendLine("");
            }

            for (int index = 0; index < this.content.Count; index++)
            {
                // Render branch start
                CaseBlockSegment segment = this.content[index];
                if (segment is CaseBlockExpressionSegment)
                {
                    output.Append(indenter.Increase().Indent);
                    output.Append("CASE ");
                    List<Expression> values = ((CaseBlockExpressionSegment)segment).Values;
                    for (int indexValue = 0; indexValue < values.Count; indexValue++)
                    {
                        Expression statement = values[indexValue];
                        output.Append(statement.GenerateBaseSource(new NullIndenter()));
                        if (indexValue < (values.Count - 1))
                            output.Append(", ");
                    }
                    output.AppendLine("");
                }
                else
                    output.AppendLine(indenter.Increase().Indent + "CASE ELSE");

                // Render branch content
                foreach (ICodeBlock statement in segment.Statements)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase().Increase()));
            }

            output.Append(indenter.Indent + "END SELECT");
            return output.ToString();
        }
    }
}
