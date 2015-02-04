using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    /// <summary>
    /// A CallSetExpressionSegment will contain one of more of these. The first one will always have at least one MemberAccessTokens but subsequent items in the
    /// set may not (eg. "a(0)(1)" is represented by segments for "a(0)" and "(1)", the second segment has no member access token).
    /// </summary>
    public class CallSetItemExpressionSegment : IExpressionSegment
    {
        private static IEnumerable<Type> AllowableTypes = new[] {
            // Note that StringToken and BuiltInValueToken are acceptable here are call target types since - although they will fail in many cases at runtime
            // (eg. "1"() or "1".a member accesses against a string will fail), they need to be allowed to fail AT RUNTIME. Numeric literals, however, would
            // result in compile errors in VBScript and so do not need to be allowed here (BuiltInValueToken includes True and False as well as values such
            // as vbObjectError).
            typeof(BuiltInFunctionToken),
            typeof(BuiltInValueToken),
            typeof(KeyWordToken),
            typeof(NameToken),
            typeof(StringToken)
        };

		public CallSetItemExpressionSegment(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments, ArgumentBracketPresenceOptions? zeroArgumentBracketsPresence)
        {
            if (memberAccessTokens == null)
                throw new ArgumentNullException("memberAccessTokens");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            MemberAccessTokens = memberAccessTokens.ToList().AsReadOnly();
            if (MemberAccessTokens.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in memberAccessTokens set");
            if (MemberAccessTokens.Any(t => t is MemberAccessorOrDecimalPointToken))
                throw new ArgumentException("MemberAccessorOrDecimalPointToken tokens should not be included in the memberAccessTokens, they are implicit as token separators");
			var firstUnacceptableTokenIfAny = MemberAccessTokens.FirstOrDefault(token => !AllowableTypes.Any(allowedType => allowedType.IsInstanceOfType(token)));
            if (firstUnacceptableTokenIfAny != null)
                throw new ArgumentException("Unacceptable token type encountered (" + firstUnacceptableTokenIfAny.GetType() + "), only allowed types are " + string.Join<Type>(", ", AllowableTypes));

            Arguments = arguments.ToList().AsReadOnly();
            if (Arguments.Any(e => e == null))
                throw new ArgumentException("Null reference encountered in arguments set");

			if (Arguments.Any())
			{
				if (zeroArgumentBracketsPresence != null)
					throw new ArgumentException("ZeroArgumentBracketsPresence must be null if there are arguments for this CallExpressionSegment");
			}
			else if (zeroArgumentBracketsPresence == null)
				throw new ArgumentException("ZeroArgumentBracketsPresence must not be null if there are zero arguments for this CallExpressionSegment");
			else if (!Enum.IsDefined(typeof(ArgumentBracketPresenceOptions), zeroArgumentBracketsPresence.Value))
				throw new ArgumentOutOfRangeException("zeroArgumentBracketsPresence");

			ZeroArgumentBracketsPresence = zeroArgumentBracketsPresence;
        }

		public enum ArgumentBracketPresenceOptions
		{
			Absent,
			Present
		}

        /// <summary>
        /// This will never be null or contain any null references. There should be considered to be implicit MemberAccessorPointTokens between each token
        /// here, if any (this will never contain any MemberAccessorOrDecimalPointToken references). The only token types that may be present in this data
        /// are BuiltInFunctionToken, BuiltInValueToken, KeyWordToken and NameToken.
        /// </summary>
        public IEnumerable<IToken> MemberAccessTokens { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references
        /// </summary>
        public IEnumerable<Expression> Arguments { get; private set; }

		/// <summary>
		/// In very particular scenarios, VBScript uses brackets to determine whether a zero-argument call is a method call or a value assignment (when
		/// setting the return value for a function, for example). This value will be null if there are any arguments and non-null if there are zero
		/// argument.s
		/// </summary>
		public ArgumentBracketPresenceOptions? ZeroArgumentBracketsPresence { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		IEnumerable<IToken> IExpressionSegment.AllTokens
		{
			get
			{
				var combinedTokens = new List<IToken>();
				var tokens = MemberAccessTokens.ToArray();
				for (var index = 0; index < tokens.Length; index++)
				{
					combinedTokens.Add(tokens[index]);
					if (index < (tokens.Length - 1))
						combinedTokens.Add(new MemberAccessorToken(combinedTokens.Last().LineIndex));
				}
				if (Arguments.Any())
				{
                    int lineIndex;
                    if (combinedTokens.Any())
                        lineIndex = combinedTokens.Last().LineIndex;
                    else
                        lineIndex = Arguments.First().AllTokens.First().LineIndex;
                    combinedTokens.Add(new OpenBrace(lineIndex));
					var arguments = Arguments.ToArray();
					for (var index = 0; index < arguments.Length; index++)
					{
						combinedTokens.AddRange(arguments[index].AllTokens);
						if (index < (arguments.Length - 1))
                            combinedTokens.Add(new ArgumentSeparatorToken(",", lineIndex));
					}
                    combinedTokens.Add(new CloseBrace(lineIndex));
				}
				return combinedTokens;
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
