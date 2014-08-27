using CSharpWriter.CodeTranslation;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.StatementTranslation
{
    public class ValueSettingStatementTranslatorTests
	{
        [Fact]
        public void UndeclaredSimpleValueTypeUpdate()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken(1, 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_env.a = 1",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0) })
            );
            var scopeAccessInformation = GetEmptyScopeAccessInformation();
            var actual = GetDefaultValueSettingStatementTranslator().Translate(
                new ValueSettingStatement(
                    expressionToSet,
                    expressionToSetTo,
                    ValueSettingStatement.ValueSetTypeOptions.Let
                ),
                scopeAccessInformation
            );
            Assert.Equal(expected, actual, new TranslatedStatementContentDetailsComparer());
        }

        [Fact]
        public void OutmostScopeDeclaredSimpleValueTypeUpdate()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken(1, 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_outer.a = 1",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0) })
            );
            var scopeAccessInformation = AddOutermostScopeVariable(
                GetEmptyScopeAccessInformation(),
                "a",
                0
            );
            var actual = GetDefaultValueSettingStatementTranslator().Translate(
                new ValueSettingStatement(
                    expressionToSet,
                    expressionToSetTo,
                    ValueSettingStatement.ValueSetTypeOptions.Let
                ),
                scopeAccessInformation
            );
            Assert.Equal(expected, actual, new TranslatedStatementContentDetailsComparer());
        }

        /// <summary>
        /// If "a" is declared as a variable in "a(0) = 1" then access will be attempted as an array or default indexed function or property
        /// </summary>
        [Fact]
        public void OutmostScopeDeclaredSimpleValueTypeUpdateOfArray()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0),
                new OpenBrace(0),
                new NumericValueToken(1, 0),
                new CloseBrace(0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken(1, 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_.SET(1, _outer.a, null, _.ARGS.Val(1))",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0) })
            );
            var scopeAccessInformation = AddOutermostScopeVariable(
                GetEmptyScopeAccessInformation(),
                "a",
                0
            );
            var actual = GetDefaultValueSettingStatementTranslator().Translate(
                new ValueSettingStatement(
                    expressionToSet,
                    expressionToSetTo,
                    ValueSettingStatement.ValueSetTypeOptions.Let
                ),
                scopeAccessInformation
            );
            Assert.Equal(expected, actual, new TranslatedStatementContentDetailsComparer());
        }

        /// <summary>
        /// If "a" is undeclared then it is implicitly treated as a variable (so this is very similar to OutmostScopeDeclaredSimpleValueTypeUpdateOfArray)
        /// </summary>
        [Fact]
        public void UndeclaredSimpleValueTypeUpdateOfArray()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0),
                new OpenBrace(0),
                new NumericValueToken(1, 0),
                new CloseBrace(0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken(1, 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_.SET(1, _env.a, null, _.ARGS.Val(1))",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0) })
            );
            var scopeAccessInformation = GetEmptyScopeAccessInformation();
            var actual = GetDefaultValueSettingStatementTranslator().Translate(
                new ValueSettingStatement(
                    expressionToSet,
                    expressionToSetTo,
                    ValueSettingStatement.ValueSetTypeOptions.Let
                ),
                scopeAccessInformation
            );
            Assert.Equal(expected, actual, new TranslatedStatementContentDetailsComparer());
        }

        /// <summary>
        /// If "a" is a function then special handling is required for "a(0) = 1"; it must compile but fail at run time
        /// </summary>
        [Fact]
        public void InvalidFunctionSettingMustCompileThoughFailAtRunTime()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0),
                new OpenBrace(0),
                new NumericValueToken(1, 0),
                new CloseBrace(0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken(1, 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_.SET(1, _.CALL(_outer, \"a\"), null, _.ARGS.Val(1))",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0) })
            );
            var scopeAccessInformation = AddOutermostScopeFunction(
                GetEmptyScopeAccessInformation(),
                "a",
                0
            );
            var actual = GetDefaultValueSettingStatementTranslator().Translate(
                new ValueSettingStatement(
                    expressionToSet,
                    expressionToSetTo,
                    ValueSettingStatement.ValueSetTypeOptions.Let
                ),
                scopeAccessInformation
            );
            Assert.Equal(expected, actual, new TranslatedStatementContentDetailsComparer());
        }

        /// <summary>
        /// This will return an empty ScopeAccessInformation that indicates an outermost scope without any statements - this does not describe a real scenario
        /// but allows us to set up data to exercise the code that the tests here are targetting
        /// </summary>
        private static ScopeAccessInformation GetEmptyScopeAccessInformation()
        {
            return ScopeAccessInformation.FromOutermostScope(
                new OutermostScope(
                    new CSharpName("UnitTestOutermostScope"),
                    new NonNullImmutableList<VBScriptTranslator.LegacyParser.CodeBlocks.ICodeBlock>()
                ),
                new NonNullImmutableList<NameToken>()
            );
        }

        private static ValueSettingStatementsTranslator GetDefaultValueSettingStatementTranslator()
		{
            return new ValueSettingStatementsTranslator(
                DefaultSupportRefName,
                DefaultSupportEnvName,
                DefaultOuterScopeName,
                DefaultNameRewriter,
    			new StatementTranslator(
                    DefaultSupportRefName,
                    DefaultSupportEnvName,
                    DefaultOuterScopeName,
                    DefaultNameRewriter,
				    GetDefaultTempValueNameGenerator(),
                    new NullLogger()
                ),
                new NullLogger()
			);
		}

        private static ScopeAccessInformation AddOutermostScopeFunction(ScopeAccessInformation scopeAccessInformation, string name, int lineIndex)
        {
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex");

            return new ScopeAccessInformation(
                scopeAccessInformation.Parent,
                scopeAccessInformation.ScopeDefiningParent,
                scopeAccessInformation.ParentReturnValueNameIfAny,
                scopeAccessInformation.ErrorRegistrationTokenIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions.Add(new ScopedNameToken(
                    name,
                    lineIndex,
                    VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions.OutermostScope
                )),
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables
            );
        }

        private static ScopeAccessInformation AddOutermostScopeVariable(ScopeAccessInformation scopeAccessInformation, string name, int lineIndex)
        {
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex");

            return new ScopeAccessInformation(
                scopeAccessInformation.Parent,
                scopeAccessInformation.ScopeDefiningParent,
                scopeAccessInformation.ParentReturnValueNameIfAny,
                scopeAccessInformation.ErrorRegistrationTokenIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions,
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables.Add(new ScopedNameToken(
                    name,
                    lineIndex,
                    VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions.OutermostScope
                ))
            );
        }

        private static CSharpName DefaultSupportRefName = new CSharpName("_");
        private static CSharpName DefaultSupportEnvName = new CSharpName("_env");
        private static CSharpName DefaultOuterScopeName = new CSharpName("_outer");
        private static VBScriptNameRewriter DefaultNameRewriter = nameToken => new CSharpName(nameToken.Content.ToLower());
		private static TempValueNameGenerator GetDefaultTempValueNameGenerator()
		{
			var index = 0;
			return (optionalPrefix, scopeAccessInformation) =>
			{
				var name = ((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + index;
				index++;
				return new CSharpName(name);
			};
		}
	}
}
