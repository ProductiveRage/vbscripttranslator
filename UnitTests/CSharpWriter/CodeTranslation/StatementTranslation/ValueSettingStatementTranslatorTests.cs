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
                new NumericValueToken("1", 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_env.a = (Int16)1",
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
        public void OutermostScopeDeclaredSimpleValueTypeUpdate()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken("1", 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_outer.a = (Int16)1",
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
        public void OutermostScopeDeclaredSimpleValueTypeUpdateOfArray()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0),
                new OpenBrace(0),
                new NumericValueToken("1", 0),
                new CloseBrace(0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken("1", 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_.SET((Int16)1, _outer.a, null, _.ARGS.Val((Int16)1))",
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
        /// If "a" is undeclared then it is implicitly treated as a variable (so this is very similar to OutermostScopeDeclaredSimpleValueTypeUpdateOfArray)
        /// </summary>
        [Fact]
        public void UndeclaredSimpleValueTypeUpdateOfArray()
        {
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0),
                new OpenBrace(0),
                new NumericValueToken("1", 0),
                new CloseBrace(0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken("1", 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_.SET((Int16)1, _env.a, null, _.ARGS.Val((Int16)1))",
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
                new NumericValueToken("1", 0),
                new CloseBrace(0)
			});
            var expressionToSetTo = new Expression(new[]
			{
                new NumericValueToken("1", 0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_.SET((Int16)1, _.CALL(_outer, \"a\"), null, _.ARGS.Val((Int16)1))",
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

        [Fact]
        public void BuiltInFunctionsAreMappedToTheSupportClassAndMayBeCalledDirectlyIfArgumentCountsMatch()
        {
            // CDate(..) needs to be mapped to _.CDATE(..) - this may be called directly if the correct number of arguments are specified. If an incorrect number
            // of arguments is passed then the support function must be executed via the "CALL" method (so that the error arises at runtime, rather than compile
            // time, in order to be consistent with VBScript), see BuiltInFunctionsAreMappedToTheSupportClassButMayNotBeCalledDirectlyIfArgumentCountsMatch.
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0)
			});
            var expressionToSetTo = new Expression(new IToken[]
			{
                new BuiltInFunctionToken("CDate", 0),
                new OpenBrace(0),
                new NameToken("a", 0),
                new CloseBrace(0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_env.a = _.CDATE(_env.a)",
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
        public void BuiltInFunctionsAreMappedToTheSupportClassButMayNotBeCalledDirectlyIfArgumentCountsMatch()
        {
            // This is a complement to BuiltInFunctionsAreMappedToTheSupportClassAndMayBeCalledDirectlyIfArgumentCountsMatch, where an incorrect number of
            // arguments is being passed to a support function. As such, it may not be called directly and must pass through the "CALL" method, so that the
            // mistake becomes a runtime error rather than compile time. On the plus side, all of the support functions may be called with ByVal parameters,
            // so the translated code is slightly more succinct that it would be if they had to support ByRef.
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0)
			});
            var expressionToSetTo = new Expression(new IToken[]
			{
                new BuiltInFunctionToken("CDate", 0),
                new OpenBrace(0),
                new NameToken("a", 0),
                new ArgumentSeparatorToken(0),
                new NameToken("b", 0),
                new CloseBrace(0)
			});
            var expected = new TranslatedStatementContentDetails(
                "_env.a = _.VAL(_.CALL(_, \"CDATE\", _.ARGS.Val(_env.a).Val(_env.b)))",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0), new NameToken("b", 0) })
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
        public void UndeclaredSetTargetsWithinFunctionsAreScopeRestrictedToThatFunction()
        {
            // The ValueSettingStatementsTranslator wasn't using the ScopeAccessInformation's GetNameOfTargetContainerIfAnyRequired extension method and
            // was incorrectly applying the logic that it should have gotten for free by using that method - if an undeclared variable was being accessed
            // within a method (for the to-set target) then it was being mapped back to the "Environment References" class instead of being treated as
            // local to the function.
            var expressionToSet = new Expression(new IToken[]
			{
                new NameToken("a", 0)
			});
            var expressionToSetTo = new Expression(new IToken[]
			{
                new NumericValueToken("1", 0)
			});
            var valueSettingStatement = new ValueSettingStatement(
                expressionToSet,
                expressionToSetTo,
                ValueSettingStatement.ValueSetTypeOptions.Let
            );

            var containingFunction = new FunctionBlock(
                isPublic: true,
                isDefault: false,
                name: new NameToken("F1", 0),
                parameters: new AbstractFunctionBlock.Parameter[0],
                statements: new[] { valueSettingStatement }
            );

            var expected = new TranslatedStatementContentDetails(
                "a = (Int16)1",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("a", 0) })
            );
            var scopeAccessInformation = GetEmptyScopeAccessInformation();
            scopeAccessInformation = new ScopeAccessInformation(
                containingFunction, // parent
                containingFunction, // scopeDefiningParent
                new CSharpName("F1"), // parentReturnValueName
                scopeAccessInformation.ErrorRegistrationTokenIfAny,
                scopeAccessInformation.DirectedWithReferenceIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions.Add(new ScopedNameToken("F1", 0, ScopeLocationOptions.WithinFunctionOrPropertyOrWith)),
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables,
                scopeAccessInformation.StructureExitPoints
            );
            var actual = GetDefaultValueSettingStatementTranslator().Translate(valueSettingStatement, scopeAccessInformation);
            Assert.Equal(expected, actual, new TranslatedStatementContentDetailsComparer());
        }

        /// <summary>
        /// This will return an empty ScopeAccessInformation that indicates an outermost scope without any statements - this does not describe a real scenario
        /// but allows us to set up data to exercise the code that the tests here are targetting
        /// </summary>
        private static ScopeAccessInformation GetEmptyScopeAccessInformation()
        {
            return ScopeAccessInformation.FromOutermostScope(
                new CSharpName("UnitTestOutermostScope"),
                new NonNullImmutableList<VBScriptTranslator.LegacyParser.CodeBlocks.ICodeBlock>(),
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
                scopeAccessInformation.DirectedWithReferenceIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions.Add(new ScopedNameToken(
                    name,
                    lineIndex,
                    VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions.OutermostScope
                )),
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables,
                scopeAccessInformation.StructureExitPoints
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
                scopeAccessInformation.DirectedWithReferenceIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions,
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables.Add(new ScopedNameToken(
                    name,
                    lineIndex,
                    VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions.OutermostScope
                )),
                scopeAccessInformation.StructureExitPoints
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
