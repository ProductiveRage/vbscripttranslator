using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public interface IExpressionSegment
    {
        string RenderedContent { get; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		IEnumerable<IToken> AllTokens { get; }
    }
}
