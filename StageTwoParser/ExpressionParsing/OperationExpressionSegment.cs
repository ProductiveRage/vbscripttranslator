using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class OperationExpressionSegment : IExpressionSegment
    {
        public OperationExpressionSegment(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (!(token is OperatorToken))
                throw new ArgumentException("The specified token must be an OperatorToken");

            Token = token;
        }

        /// <summary>
        /// This will never be null and it will always be an OperatorToken
        /// </summary>
        public IToken Token { get; private set; }

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
