using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class BracketedExpressionSegment : IExpressionSegment
    {
        private readonly ReadOnlyCollection<IToken> _allTokens;
        public BracketedExpressionSegment(IEnumerable<IExpressionSegment> segments)
        {
			if (segments == null)
                throw new ArgumentNullException("segments");

			Segments = segments.ToList().AsReadOnly();
			if (Segments.Any(e => e == null))
				throw new ArgumentException("Null reference encountered in segments set");
			if (!Segments.Any())
				throw new ArgumentException("Empty segments set specified - invalid");

            // 2015-03-23 DWR: For deeply-nested bracketed segments, it can be very expensive to enumerate over their AllTokens sets repeatedly so it's worth preparing the data once and
            // avoiding doing it over and over again. This is often seen with an expression with many string concatenations - currently they are broken down into pairs of operations,
            // which results in many bracketed operations (I want to change this for concatenations going forward, since it's so common to have sets of concatenations and it would
            // be better if the CONCAT took a variable number of arguments rather than just two, but this hasn't been done yet).
            _allTokens =
                new IToken[] { new OpenBrace(Segments.First().AllTokens.First().LineIndex) }
                .Concat(Segments.SelectMany(s => s.AllTokens))
                .Concat(new[] { new CloseBrace(Segments.Last().AllTokens.Last().LineIndex) })
                .ToList()
                .AsReadOnly();
		}

        /// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		public IEnumerable<IExpressionSegment> Segments { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
        IEnumerable<IToken> IExpressionSegment.AllTokens { get { return _allTokens; } }

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
