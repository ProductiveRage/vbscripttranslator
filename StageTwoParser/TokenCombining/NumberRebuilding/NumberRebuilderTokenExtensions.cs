using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding
{
    public static class NumberRebuilderTokenExtensions
    {
        public static bool Is<T>(this IToken token) where T : IToken
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return (token.GetType() == typeof(T));
        }

        public static bool CouldPrecedeDecimalPointOrNegativeSign(this IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return (token.Is<OpenBrace>() || token.Is<OperatorToken>());
        }

        public static bool IsMinusSignOperator(this IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return token.Is<OperatorToken>() && (token.Content == "-");
        }
    }
}
