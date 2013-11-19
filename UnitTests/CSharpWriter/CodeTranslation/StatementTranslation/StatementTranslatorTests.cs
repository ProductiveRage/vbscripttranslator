using CSharpWriter.CodeTranslation;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.StatementTranslation
{
	public class StatementTranslatorTests
	{
		// TODO: "o" where "o" has a default parameter-less property => try to access that property, doesn't matter if returns value-type or reference
		// TODO: "o" where "o" has a default parameter-less function => try to access that property, doesn't matter if returns value-type or reference

		// TODO: "WScript.Echo o" where "o" has a default parameter-less property => displays the property value (if value-type, "Type mismatch" if reference)
		// TODO: "WScript.Echo o" where "o" has a default parameter-less function => "Type mismatch"
		// TODO: "WScript.Echo o()" where "o" has a default parameter-less function => displays the function return value (if value-type, "Type mismatch" if reference)

		[Fact]
		public void IsolatedNonFunctionOrPropertyReferenceHasValueTypeAccessLogic()
		{
			// "o" (where there is no function or property in scope called "o")
			var expression = new Expression(new[]
			{
				new CallExpressionSegment(
					new[] { new NameToken("o", 0) },
					new Expression[0],
					CallExpressionSegment.ArgumentBracketPresenceOptions.Absent
				)
			});
			var expected = new TranslatedStatementContentDetails(
				"_.VAL(o)",
				new NonNullImmutableList<NameToken>(new[] { new NameToken("o", 0) })
			);
			var scopeAccessInformation = ScopeAccessInformation.Empty;
			Assert.Equal(
				expected,
				GetDefaultStatementTranslator().Translate(expression, scopeAccessInformation, ExpressionReturnTypeOptions.None),
				new TranslatedStatementContentDetailsComparer()
			);
		}

		[Fact]
		public void IsolatedFunctionCallAccordingToScopeDoesNotHaveValueTypeAccessLogic()
		{
			// "o" (where there is a function in scope called "o")
			var expression = new Expression(new[]
			{
				new CallExpressionSegment(
					new[] { new NameToken("o", 0) },
					new Expression[0],
					CallExpressionSegment.ArgumentBracketPresenceOptions.Absent
				)
			});
			var scopeAccessInformation = new ScopeAccessInformation(
				null,
				null,
				null,
				new NonNullImmutableList<NameToken>(),
				new NonNullImmutableList<NameToken>(new[] { new NameToken("o", 0) }), // Functions
				new NonNullImmutableList<NameToken>(),
				new NonNullImmutableList<NameToken>()
			);
			var expected = new TranslatedStatementContentDetails(
				"o()",
				new NonNullImmutableList<NameToken>(new[] { new NameToken("o", 0) })
			);
			Assert.Equal(
				expected,
				GetDefaultStatementTranslator().Translate(expression, scopeAccessInformation, ExpressionReturnTypeOptions.None),
				new TranslatedStatementContentDetailsComparer()
			);
		}

		private static StatementTranslator GetDefaultStatementTranslator()
		{
			return new StatementTranslator(
                DefaultSupportClassName,
                DefaultSupportEnvName,
                DefaultNameRewriter,
				GetDefaultTempValueNameGenerator(),
                new NullLogger()
			);
		}

        private static CSharpName DefaultSupportClassName = new CSharpName("_");
        private static CSharpName DefaultSupportEnvName = new CSharpName("__");
        private static VBScriptNameRewriter DefaultNameRewriter = nameToken => new CSharpName(nameToken.Content.ToLower());
		private static TempValueNameGenerator GetDefaultTempValueNameGenerator()
		{
			var index = 0;
			return optionalPrefix =>
			{
				var name = optionalPrefix.Name + "_tempVal" + index;
				index++;
				return new CSharpName(name);
			};
		}
	}
}
