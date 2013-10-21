using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class BracketedExpressionSegment : IExpressionSegment
    {
        public BracketedExpressionSegment(IEnumerable<IExpressionSegment> segments)
        {
			if (segments == null)
                throw new ArgumentNullException("segments");

			Segments = segments.ToList().AsReadOnly();
			if (Segments.Any(e => e == null))
				throw new ArgumentException("Null reference encountered in segments set");
			if (!Segments.Any())
				throw new ArgumentException("Empty segments set specified - invalid");
		}

        /// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		public IEnumerable<IExpressionSegment> Segments { get; private set; }

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
				tokens.AddRange(Segments.SelectMany(s => s.AllTokens));
				tokens.Add(new CloseBrace(")"));
				return tokens;
			}
		}

        public string RenderedContent
        {
            get
            {
                return "(" + string.Join("", Segments.Select(e => e.RenderedContent)) + ")";
            }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
