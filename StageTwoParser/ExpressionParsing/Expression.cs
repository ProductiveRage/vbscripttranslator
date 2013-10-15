using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class Expression
    {
        public Expression(IEnumerable<IExpressionSegment> segments)
        {
            if (segments == null)
                throw new ArgumentNullException("segments");

            Segments = segments.ToList().AsReadOnly();
            if (!Segments.Any())
                throw new ArgumentException("The segments set may not be empty");
            if (Segments.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in segments set");
        }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        public IEnumerable<IExpressionSegment> Segments { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		public IEnumerable<IToken> AllTokens
		{
			get { return Segments.SelectMany(s => s.AllTokens); }
		}

        public string RenderedContent
        {
            get
            {
                return string.Join(
                    "",
                    Segments.Select(s => s.RenderedContent)
                );
            }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
