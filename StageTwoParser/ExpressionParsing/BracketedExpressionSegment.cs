using System;
using System.Collections.Generic;
using System.Linq;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class BracketedExpressionSegment : IExpressionSegment
    {
        public BracketedExpressionSegment(IEnumerable<Expression> expressions)
        {
            if (expressions == null)
                throw new ArgumentNullException("segments");

            Expressions = expressions.ToList().AsReadOnly();
            if (Expressions.Any(e => e == null))
                throw new ArgumentException("Null reference encountered in expressions set");
        }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        public IEnumerable<Expression> Expressions { get; private set; }

        public string RenderedContent
        {
            get
            {
                return "(" + string.Join("", Expressions.Select(e => e.RenderedContent)) + ")";
            }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
