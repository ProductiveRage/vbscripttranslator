using System;
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
                "_.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_env.i, v1 => { _env.i = v1; }));"
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
                "    _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(i, v2 => { i = v2; }));",
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
                "    _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(i, v2 => { i = v2; }));",
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
                "    _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(_outer.i, v2 => { _outer.i = v2; }));",
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
                "_.CALL(_env.wscript, \"echo\", _.ARGS.Val(_.CONCAT(_env.a, _env.b, _env.c, _env.d)));"
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
                "_.CALL(_env.wscript, \"echo\", _.ARGS.Val(_.CONCAT(_env.a, _.ADD((Int16)1, (Int16)2), _env.c, _env.d)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
