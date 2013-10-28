using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class BuiltInValueExpressionSegment : IExpressionSegment
    {
		public BuiltInValueExpressionSegment(BuiltInValueToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            Token = token;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
		public BuiltInValueToken Token { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		IEnumerable<IToken> IExpressionSegment.AllTokens { get { return new[] { Token }; } }

        public string RenderedContent
        {
            get { return Token.Content; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
