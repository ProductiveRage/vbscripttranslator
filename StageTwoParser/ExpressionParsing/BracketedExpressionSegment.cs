using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

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
			if (!Expressions.Any())
				throw new ArgumentException("Empty expressions set specified - invalid");
		}

        /// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
        public IEnumerable<Expression> Expressions { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		IEnumerable<IToken> IExpressionSegment.AllTokens
		{
			get
			{
				var tokens = new List<IToken>
				{
					new OpenBrace("(")
				};
				tokens.AddRange(Expressions.SelectMany(e => e.AllTokens));
				tokens.Add(new CloseBrace(")"));
				return tokens;
			}
		}

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
