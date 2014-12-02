using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class SelectBlock : IHaveNestedContent, ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private readonly List<CommentStatement> _openingComments;
        private readonly List<CaseBlockSegment> _content;
        public SelectBlock(
            Expression expression,
            IEnumerable<CommentStatement> openingComments,
            IEnumerable<CaseBlockSegment> content)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (openingComments == null)
                throw new ArgumentNullException("openingComments");
            if (content == null)
                throw new ArgumentNullException("content");

            _openingComments = openingComments.ToList();
            if (_openingComments.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in openingComments set");

            _content = content.ToList();
            if (_openingComments.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in content set");
            var firstUnsupportedContentSegment = _content.FirstOrDefault(c => !(c is CaseBlockExpressionSegment) && !(c is CaseBlockElseSegment));
            if (firstUnsupportedContentSegment != null)
                throw new ArgumentException("Unrecognised content element: " + firstUnsupportedContentSegment.GetType());
            if (_content.First() is CaseBlockElseSegment)
                throw new ArgumentException("First content element may not be a CaseBlockElseSegment");
            if (((IEnumerable<CaseBlockSegment>)_content).Reverse().Skip(1).Any(c => c is CaseBlockElseSegment))
                throw new ArgumentException("Only the last content segment may be a CaseBlockElseSegment (and only when there are multiple content segments)");

            Expression = expression;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null
        /// </summary>
        public Expression Expression { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set
        /// </summary>
        public IEnumerable<CommentStatement> OpeningComments { get { return _openingComments.AsReadOnly(); } }

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set. All items will be CaseBlockExpressionSegment or
        /// CaseBlockElseSegment instances and only the last segment may be a CaseBlockElseSegment (and only if there are multiple segments).
        /// </summary>
        public IEnumerable<CaseBlockSegment> Content { get { return _content.AsReadOnly(); } }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get
            {
                return new ICodeBlock[] { Expression }
                    .Concat(_content.Select(c => c as CaseBlockExpressionSegment).Where(c => c != null).SelectMany(c => c.Values))
                    .Concat(_content.SelectMany(c => c.Statements));
            }
        }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        public abstract class CaseBlockSegment
        {
            private readonly List<ICodeBlock> _statements;
            protected CaseBlockSegment(IEnumerable<ICodeBlock> statements)
            {
                if (statements == null)
                    throw new ArgumentNullException("statements");

                _statements = statements.ToList();
                if (_statements.Any(v => v == null))
                    throw new ArgumentException("Null reference encountered in statements set");
            }

            /// <summary>
            /// This will never be null nor contain any null references, but it may be an empty set
            /// </summary>
            public IEnumerable<ICodeBlock> Statements { get { return _statements.AsReadOnly(); } }
        }

        public class CaseBlockExpressionSegment : CaseBlockSegment
        {
            private readonly List<Expression> _values;
            private readonly List<ICodeBlock> _statements;
            public CaseBlockExpressionSegment(IEnumerable<Expression> values, IEnumerable<ICodeBlock> statements) : base(statements)
            {
                if (values == null)
                    throw new ArgumentNullException("values");

                _values = values.ToList();
                if (_values.Any(v => v == null))
                    throw new ArgumentException("Null reference encountered in openingComments set");
                if (!_values.Any())
                    throw new ArgumentException("values is an empty set  - invalid");
            }

            /// <summary>
            /// This will never be null, empty nor contain any null references
            /// </summary>
            public IEnumerable<Expression> Values { get { return _values.AsReadOnly(); } }
        }
        
        public class CaseBlockElseSegment : CaseBlockSegment
        {
            public CaseBlockElseSegment(IEnumerable<ICodeBlock> statements) : base(statements) { }
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
            
            output.Append(indenter.Indent + "SELECT CASE ");
            output.AppendLine(Expression.GenerateBaseSource(new NullIndenter()));

            if (_openingComments != null)
            {
                foreach (CommentStatement statement in _openingComments)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));
                output.AppendLine("");
            }

            for (int index = 0; index < _content.Count; index++)
            {
                // Render branch start
                CaseBlockSegment segment = _content[index];
                if (segment is CaseBlockExpressionSegment)
                {
                    output.Append(indenter.Increase().Indent);
                    output.Append("CASE ");
                    var valuesArray = ((CaseBlockExpressionSegment)segment).Values.ToArray();
                    for (int indexValue = 0; indexValue < valuesArray.Length; indexValue++)
                    {
                        Expression statement = valuesArray[indexValue];
                        output.Append(statement.GenerateBaseSource(new NullIndenter()));
                        if (indexValue < (valuesArray.Length - 1))
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
