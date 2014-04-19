using CSharpWriter.CodeTranslation;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using VBScriptTranslator.LegacyParser.Tokens;
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
				"_.VAL(_env.o)",
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

            var scopeAccessInformation = AddOutermostScopeFunction(
                ScopeAccessInformation.Empty,
                "o",
                0
            );
            var expected = new TranslatedStatementContentDetails(
                "_.CALL(_outer, \"o\")",
                new NonNullImmutableList<NameToken>(new[] { new NameToken("o", 0) })
            );
            Assert.Equal(
                expected,
                GetDefaultStatementTranslator().Translate(expression, scopeAccessInformation, ExpressionReturnTypeOptions.None),
                new TranslatedStatementContentDetailsComparer()
            );
        }

        [Fact]
        public void KnownVariablePassedAsArgumentToKnownFunctionIsPassedByRef()
        {
            // "o(a)" (where there is a function in scope called "o" and a variable "a")
            var expression = new Expression(new[]
			{
				new CallExpressionSegment(
					new[] { new NameToken("o", 0) },
					new[]
                    {
                        new Expression(new[] {
                            new CallExpressionSegment(new[] { new NameToken("a", 0) }, new Expression[0], CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent)
                        })
                    },
					null
				)
			});

            var scopeAccessInformation = AddOutermostScopeVariable(
                AddOutermostScopeFunction(
                    ScopeAccessInformation.Empty,
                    "o",
                    0
                ),
                "a",
                0
            );
            var expected = new TranslatedStatementContentDetails(
                "_.CALL(_outer, \"o\", _.ARGS.Ref(_outer.a, v0 => { _outer.a = v0; }).GetArgs())",
                new NonNullImmutableList<NameToken>(new[] { 
                    new NameToken("a", 0),
                    new NameToken("o", 0)
                })
            );
            Assert.Equal(
                expected,
                GetDefaultStatementTranslator().Translate(expression, scopeAccessInformation, ExpressionReturnTypeOptions.None),
                new TranslatedStatementContentDetailsComparer()
            );
        }

        /// <summary>
        /// VBScript will give special significance to arguments that are wrapped in extra brackets - if the argument would have been passed ByRef
        /// before, the brackets will force it to be passed ByVal
        /// </summary>
        [Fact]
        public void KnownVariablePassedAsArgumentToKnownFunctionIsPassedByValIfWrappedInBrackets()
        {
            // "o((a))" (where there is a function in scope called "o" and a variable "a")
            var expression = new Expression(new[]
			{
				new CallExpressionSegment(
					new[] { new NameToken("o", 0) },
					new[]
                    {
                        new Expression(new[] {
                            new BracketedExpressionSegment(new[] {
                                new CallExpressionSegment(new[] { new NameToken("a", 0) }, new Expression[0], CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent)
                            })
                        })
                    },
					null
				)
			});

            var scopeAccessInformation = AddOutermostScopeVariable(
                AddOutermostScopeFunction(
                    ScopeAccessInformation.Empty,
                    "o",
                    0
                ),
                "a",
                0
            );
            var expected = new TranslatedStatementContentDetails(
                "_.CALL(_outer, \"o\", _.ARGS.Val(_outer.a).GetArgs())",
                new NonNullImmutableList<NameToken>(new[] { 
                    new NameToken("a", 0),
                    new NameToken("o", 0)
                })
            );
            Assert.Equal(
                expected,
                GetDefaultStatementTranslator().Translate(expression, scopeAccessInformation, ExpressionReturnTypeOptions.None),
                new TranslatedStatementContentDetailsComparer()
            );
        }

        [Fact]
        public void NestedFunctionOrArrayAccess()
        {
            // "a(0)(b)" (where neither a nor b are defined and so there could be method calls OR array accesses)
            var expression = new Expression(new[]
			{
                new CallSetExpressionSegment(new[]
                {
                    new CallSetItemExpressionSegment(
                        new[] { new NameToken("a", 0) },
                        new[] { new Expression(new[] { new NumericValueExpressionSegment(new NumericValueToken(0, 0)) }) },
                        null
                    ),
                    new CallSetItemExpressionSegment(
                        new IToken[0],
                        new[] { new Expression(new[] { new CallExpressionSegment(
                            new[] { new NameToken("b", 0) },
                            new Expression[0],
                            CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent
                        )})},
                        null
                    )
                })
			});

            // Since we can't know until runtime if "a" is an array that is being accessed or a function/property, the arguments need to
            // be constructed to work as ByVal or ByRef if it IS a function or property. Since "0" is a constant it will be ByVal but
            // since "b" is a variable it has to be marked as exligible for ByRef (this will not have any effect if "a(0)" is an
            // array or if it is an object with a default function or property whose argument is marked as ByVal, but we won't
            // know that until runtime).
            var expected = new TranslatedStatementContentDetails(
                "_.CALL(_.CALL(_env.a, _.ARGS.Val(0).GetArgs()), _.ARGS.Ref(_env.b, v0 => { _env.b = v0; }).GetArgs())",
                new NonNullImmutableList<NameToken>(new[] {
                    new NameToken("a", 0),
                    new NameToken("b", 0)
                })
            );
            Assert.Equal(
                expected,
                GetDefaultStatementTranslator().Translate(expression, ScopeAccessInformation.Empty, ExpressionReturnTypeOptions.None),
                new TranslatedStatementContentDetailsComparer()
            );
        }

        private static StatementTranslator GetDefaultStatementTranslator()
		{
			return new StatementTranslator(
                DefaultsupportRefName,
                DefaultSupportEnvName,
                DefaultOuterScopeName,
                DefaultNameRewriter,
				GetDefaultTempValueNameGenerator(),
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
                scopeAccessInformation.ParentIfAny,
                scopeAccessInformation.ScopeDefiningParentIfAny,
                scopeAccessInformation.ParentReturnValueNameIfAny,
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
                scopeAccessInformation.ParentIfAny,
                scopeAccessInformation.ScopeDefiningParentIfAny,
                scopeAccessInformation.ParentReturnValueNameIfAny,
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

        private static CSharpName DefaultsupportRefName = new CSharpName("_");
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
