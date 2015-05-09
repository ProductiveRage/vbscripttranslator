using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndMiscTranslationTests
    {
        // TODO: Test function call with numeric values (1 and 1.1), string values, built-in values and built-in functions (such as "Now") and ensure that
        // they all have the arguments specified as ByVal
        // - Is it easiest to put it in here or better to put it into StatementTranslatorTests?

        /// <summary>
        /// The code here accesses an undeclared variable in a statement in the outermost scope, that scope should be registered in the EnvironmentReferences
        /// class. There is also a "wscript" reference which is declared as an External Dependency in the translator, this will appear in the Environment
        /// References class as well (as any/all External Dependencies should).
        /// </summary>
        [Fact]
        public void UndeclaredVariablesInTheOutermostScopeShouldBeDefinedAsAnEnvironmentVariable()
        {
            var source = @"
                WScript.Echo i
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(_env.i, v1 => { _env.i = v1; }));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This code will access an undeclared variable within a function. The scope of that undeclared variable should be restricted to the function in
        /// which it is accessed and not bleed out into the outer scope.
        /// </summary>
        [Fact]
        public void UndeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
                Test1
                Function Test1()
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_.CALL(_outer, \"test1\");",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    object i = null; /* Undeclared in source */",
                "    _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(i, v2 => { i = v2; }));",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [Fact]
        public void DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
                Test1
                Function Test1()
                    Dim i
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_.CALL(_outer, \"test1\");",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    object i = null;",
                "    _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(i, v2 => { i = v2; }));",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [Fact]
        public void DeclaredVariableInOutermostScopeShouldBeAccessedFromThereWhenRequiredWithinFunction()
        {
            var source = @"
                Dim i
                Test1
                Function Test1()
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_.CALL(_outer, \"test1\");",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    _.CALL(_env.wscript, \"Echo\", _.ARGS.Ref(_outer.i, v2 => { _outer.i = v2; }));",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void NumericLiteralsAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func 1()";
            var expected = new[]
            {
                "_.CALL(_env.func, _.ARGS.Val(_.RAISEERROR(new TypeMismatchException(\"\\'[number: 1]\\\' is called like a function\"))));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void StringLiteralsAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func \"1\"()";
            var expected = new[]
            {
                "_.CALL(_env.func, _.ARGS.Val(_.RAISEERROR(new TypeMismatchException(\"\\'[string: \\\"1\\\"]\\\' is called like a function\"))));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void BuiltinValuesAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func vbObjectError()";
            var expected = new[]
            {
                "_.CALL(_env.func, _.ARGS.Val(_.RAISEERROR(new TypeMismatchException(\"\\'vbObjectError\\' is called like a function\"))));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void ClassNameFollowedByBracketsInNewStatementResultsInCompileTimeError()
        {
            var source = "c = new C1()";
            Assert.Throws<Exception>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }

        /// <summary>
        /// Since runs of string concatenations are so common, an exception to the two-arguments-per-operation (apart from NOT, that only takes one) is made
        /// to allow the values to be combined in a single CONCAT call, reducing the size of the emitted code
        /// </summary>
        [Fact]
        public void ConcatFunctionAllowsMoreThanTwoArguments()
        {
            var source = @"
                WScript.Echo a & b & c & d
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CONCAT(_env.a, _env.b, _env.c, _env.d)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is related to the ConcatFunctionAllowsMoreThanTwoArguments and provides reassurance that string concatenations will only be joined if it
        /// would have no effect on the rest of processing (since the addition operation should take precedence, there is no CONCAT-flattening that can
        /// be performed in this case)
        /// </summary>
        [Fact]
        public void ConcatFunctionAllowsMoreThanTwoArgumentsButDoesNotAffectNestedOperationsOfOtherTypes()
        {
            var source = @"
                WScript.Echo a & 1 + 2 & c & d
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CONCAT(_env.a, _.ADD((Int16)1, (Int16)2), _env.c, _env.d)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// The string values that specify target member names in a CALL expression must not be manipulated by the name rewriter at runtime. This means that
        /// their casing will not be affected and - more importantly, any manipulations relating to C# keywords will NOT be applied. When the target is a
        /// translated class, the name rewriter manipulations would not cause any issue but if the target is not something that is translated (a COM component,
        /// for example), then trying to access its members with the name-rewritten versions will fail. This means that the CALL implementation must be able to
        /// consider the same name rewriter rules at runtime that the translator does.
        /// </summary>
        [Fact]
        public void MemberAccessorsInCallStatementsShouldNotBeRenamedAtTranslationTime()
        {
            // "Params" is a C# keyword, so we couldn't emit translated code with a method called "Params", but if "a" is an external reference (such as a COM
            // component) then It may have a methor or property named "Params". As such we mustn't enforce the rewriting of "Params" to something C#-friendly
            // at compile time (the CALL implementation will have to do some magic)
            // - The GetTranslatedStatements uses the DefaultTranslator which uses the DefaultRuntimeSupportClassFactory.DefaultNameRewriter which will
            //   ensure that C# keywords are rewritten to something safe
            var source = @"
                WScript.Echo a.Params
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"Echo\", _.ARGS.Val(_.CALL(_env.a, \"Params\")));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Theory, MemberData("VariousBracketDeterminedRefValArgumentData")]
        public void VariousBracketDeterminedRefValArgumentCases(string source, string expectedResult)
        {
            var translatedContent = WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            Assert.Equal(expectedResult, translatedContent.Select(c => c.Trim()).Single(c => c != ""));
        }

        public static IEnumerable<object[]> VariousBracketDeterminedRefValArgumentData
        {
            get
            {
                yield return new object[] { "func x", "_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }));" };
                yield return new object[] { "func (x)", "_.CALL(_env.func, _.ARGS.Val(_env.x));" };

                yield return new object[] { "func x, y", "_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }).Ref(_env.y, v2 => { _env.y = v2; }));" };
                yield return new object[] { "func (x), y", "_.CALL(_env.func, _.ARGS.Val(_env.x).Ref(_env.y, v1 => { _env.y = v1; }));" };
                yield return new object[] { "func x, (y)", "_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }).Val(_env.y));" };

                yield return new object[] { "z = func(x)", "_env.z = _.VAL(_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; })));" };
                yield return new object[] { "z = func(x, y)", "_env.z = _.VAL(_.CALL(_env.func, _.ARGS.Ref(_env.x, v1 => { _env.x = v1; }).Ref(_env.y, v2 => { _env.y = v2; })));" };
                yield return new object[] { "z = func((x), y)", "_env.z = _.VAL(_.CALL(_env.func, _.ARGS.Val(_env.x).Ref(_env.y, v1 => { _env.y = v1; })));" };
            }
        }
    }
}
