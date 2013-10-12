using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser
{
    public class Expression
    {
        public Expression(IEnumerable<ExpressionSegment> segments)
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
        public IEnumerable<ExpressionSegment> Segments { get; private set; }
    }
}
