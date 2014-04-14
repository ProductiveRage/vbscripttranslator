using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    /// <summary>
    /// A standalone CallExpressionSegment is essentially a specialised version of the CallSetItemExpressionSegment where there must be at least one Member
    /// Access Token.
    /// </summary>
    public class CallExpressionSegment : CallSetItemExpressionSegment
    {
        public CallExpressionSegment(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments, ArgumentBracketPresenceOptions? zeroArgumentBracketsPresence)
            : base(memberAccessTokens, arguments, zeroArgumentBracketsPresence)
        {
            if (!base.MemberAccessTokens.Any())
                throw new ArgumentException("The memberAccessTokens set may not be empty");
        }

        /// <summary>
        /// This will never be null, empty or contain any null references. There should be considered to be implicit MemberAccessorPointTokens between each
        /// token here (this will never contain any MemberAccessorOrDecimalPointToken references). The only token types that may be present in this data are
        /// BuiltInFunctionToken, BuiltInValueToken, KeyWordToken and NameToken.
        /// </summary>
        public new IEnumerable<IToken> MemberAccessTokens { get { return base.MemberAccessTokens; } }
    }
}
