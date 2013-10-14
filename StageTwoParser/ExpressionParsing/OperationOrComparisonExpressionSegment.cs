using System;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class OperatorOrComparisonExpressionSegment : IExpressionSegment
    {
        public OperatorOrComparisonExpressionSegment(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (!(token is OperatorToken) && !(token is ComparisonToken))
                throw new ArgumentException("The specified token must be either an OperatorToken or a ComparisonToken");

            Token = token;
        }

        /// <summary>
        /// This will never be null and it will always be either an OperatorToken or a ComparisonToken
        /// </summary>
        public IToken Token { get; private set; }

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
