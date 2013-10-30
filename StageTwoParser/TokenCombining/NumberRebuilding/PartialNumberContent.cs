using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding
{
    public class PartialNumberContent
    {
        public PartialNumberContent(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            Tokens = tokens.ToList().AsReadOnly();
            if (Tokens.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in tokens set");
            if (Tokens.Any(t => !IsValidToken(t)))
                throw new ArgumentException("The only allowable tokens are minus sign OperatorTokens, numeric AtomTokens and MemberAccessorOrDecimalPointTokens");
        }
        public PartialNumberContent() : this(new IToken[0]) { }

        /// <summary>
        /// This will never be null nor contain any null references. Any tokens in the - TODO: What 
        /// </summary>
        public IEnumerable<IToken> Tokens { get; private set; }

        /// <summary>
        /// If this returns a number then all of the tokens represented by this instance will be used to describe it. If this returns null then this was only
        /// a partial attempt at constructing a number and, if there are are no more tokens to process, then all of the tokens may be pulled back out as
        /// unprocessed
        /// </summary>
        public double? TryToExpressNumberFromTokens()
        {
            double value;
            if (double.TryParse(string.Join("", Tokens.Select(t => t.Content)), out value))
                return value;
            return null;
        }

        /// <summary>
        /// This will return a new instance with a new token set consisting of the current instance's tokens and that specified by the token argument
        /// </summary>
        public PartialNumberContent AddToken(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (!IsValidToken(token))
                throw new ArgumentException("The only allowable tokens are minus sign OperatorTokens, numeric AtomTokens and MemberAccessorOrDecimalPointTokens");

            return new PartialNumberContent(Tokens.Concat(new[] { token }));
        }

        private static bool IsValidToken(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return token.IsMinusSignOperator() || (token is NumericValueToken) || token.Is<MemberAccessorOrDecimalPointToken>();
        }
    }
}
