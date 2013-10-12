using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser
{
    public class ExpressionSegment
    {
        public ExpressionSegment(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments)
        {
            if (memberAccessTokens == null)
                throw new ArgumentNullException("memberAccessTokens");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            MemberAccessTokens = memberAccessTokens.ToList().AsReadOnly();
            if (!MemberAccessTokens.Any())
                throw new ArgumentException("The memberAccessTokens set may not be empty");
            if (MemberAccessTokens.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in memberAccessTokens set");
            if (MemberAccessTokens.Any(t => t is MemberAccessorOrDecimalPointToken))
                throw new ArgumentException("MemberAccessorOrDecimalPointToken tokens should not be included in the memberAccessTokens, they are implicit as token separators");

            Arguments = arguments.ToList().AsReadOnly();
            if (Arguments.Any(e => e == null))
                throw new ArgumentException("Null reference encountered in arguments set");
        }

        /// <summary>
        /// This will never be null, empty or contain any null references. There should be considered to be implicit MemberAccessorPointTokens
        /// between each token here (this will never contain any MemberAccessorOrDecimalPointToken references).
        /// </summary>
        public IEnumerable<IToken> MemberAccessTokens { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references
        /// </summary>
        public IEnumerable<Expression> Arguments { get; private set; }
    }
}
