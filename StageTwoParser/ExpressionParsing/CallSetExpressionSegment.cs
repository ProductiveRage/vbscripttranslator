using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    /// <summary>
    /// Consecutive CallExpressionSegment instance should not exist within an Expression if they all relates to parts of the same operation, instead
    /// they should be wrapped in a CallSetExpressionSegment so that it is clear that they are part of one action. This is because "a(0).Test" should
    /// be seen as one retrieval and not the constituent parts "a(0)" and "Test".
    /// </summary>
    public class CallSetExpressionSegment : IExpressionSegment
    {
        public CallSetExpressionSegment(IEnumerable<CallSetItemExpressionSegment> callExpressionSegments)
        {
            if (callExpressionSegments == null)
                throw new ArgumentNullException("callExpressionSegments");

            CallExpressionSegments = callExpressionSegments.ToList().AsReadOnly();
            if (CallExpressionSegments.Count() < 2)
                throw new ArgumentException("The callExpressionSegments set must have at least two segments otherwise a single CallExpressionSegment should descrive this data");
            if (CallExpressionSegments.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in callExpressionSegments set");
        }

        /// <summary>
        /// This will never be null, empty (it will always have at least two entries otherwise a single CallExpressionSegment would be appropriate
        /// rather than a CallSetExpressionSegment) or contain any null references. The first segment will always have at least Member Access
        /// Token, thbough subsequent ones may not (to represent, for example, "a(0)(1)" - there would be segments for "a(0)" and "(1)",
        /// the second segment would have no Member Access Tokens).
        /// </summary>
        public IEnumerable<CallSetItemExpressionSegment> CallExpressionSegments { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		IEnumerable<IToken> IExpressionSegment.AllTokens
		{
			get
			{
                var tokens = new List<IToken>();
                var numberOfExpressions = CallExpressionSegments.Count();
                for (var index = 0; index < numberOfExpressions; index++)
                {
                    tokens.AddRange(((IExpressionSegment)CallExpressionSegments.ElementAt(index)).AllTokens);
                    if (index < (numberOfExpressions - 1))
                        tokens.Add(new MemberAccessorToken(tokens.Last().LineIndex));
                }
                return tokens;
			}
		}

        public string RenderedContent
        {
            get
            {
				return string.Join(
					"",
					((IExpressionSegment)this).AllTokens.Select(t =>
						(t is StringToken) ? ("\"" + t.Content.Replace("\"", "\"\"") + "\"") : t.Content
					)
				);
            }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
