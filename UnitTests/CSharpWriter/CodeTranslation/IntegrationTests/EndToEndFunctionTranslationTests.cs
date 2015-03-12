using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndFunctionTranslationTests
    {
        /// <summary>
        /// When a function (or property) has only a single executable statement that is a return-this-expression statement, then this can be translated
        /// into a single line C# return statement. Anything more complicated requires a temporary variable which is used to track the return value and
        /// returned from any exit point.
        /// </summary>
        [Fact]
        public void IfTheOnlyExecutableStatementIsReturnValueThenTranslateIntoSingleReturnStatement()
        {
            var source = @"
                PUBLIC FUNCTION F1()
                    ' Test simple-return-format functions
                    F1 = CDate(""2007-04-01"")
                END FUNCTION
            ";
            var expected = new[]
            {
                "public object f1()",
                "{",
                "    // Test simple-return-format functions",
                "    return _.VAL(_.CDATE(\"2007-04-01\"));",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is very similar to IfTheOnlyExecutableStatementIsReturnValueThenTranslateIntoSingleReturnStatement except that it demonstrates the difference
        /// required when the function return value is SET - meaning that it must be an object reference (however, because it references an undeclared variable,
        /// that variable must be defined within the function scope; so the C# is no longer a one-executable-line job, but the principle remains).
        /// </summary>
        [Fact]
        public void IfTheOnlyExecutableStatementIsSetReturnValueThenTranslateIntoSingleReturnStatement()
        {
            var source = @"
                PUBLIC FUNCTION F1()
                    Set F1 = a
                END FUNCTION
            ";
            var expected = new[]
            {
                "public object f1()",
                "{",
                "    object a = null; /* Undeclared in source */",
                "    return _.OBJ(a);",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is very similar to IfTheOnlyExecutableStatementIsSetReturnValueThenTranslateIntoSingleReturnStatement except that it demonstrates the fact that
        /// if the return reference is already known to be an object type (which "Nothing" is) then it doesn't need to call the OBJ method for safety.
        /// </summary>
        [Fact]
        public void IfTheOnlyExecutableStatementIsSetKnownObjectReturnValueThenTranslateIntoSingleReturnStatement()
        {
            var source = @"
                PUBLIC FUNCTION F1()
                    Set F1 = Nothing
                END FUNCTION
            ";
            var expected = new[]
            {
                "public object f1()",
                "{",
                "    return VBScriptConstants.Nothing;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
        /// <summary>
        /// If a ByRef argument of a function is passed into another function as a ByRef argument then it must be stored in a temporary variable and then
        /// updated from this variable after the function call completes (whether it succeeds or fails - if the ByRef argument was altered before the error
        /// then that updated value must be persisted). This is to avoid trying to access "ref" reference in a lambda, which is a compile error in C#.
        /// </summary>
        [Fact]
        public void ByRefFunctionArgumentRequiresSpecialTreatmentIfUsedElsewhereAsByRefArgument()
        {
            var source = @"
                Function F1(a)
                    F2 a
                End Function

                Function F2(a)
                End Function
            ";
            var expected = new[]
            {
                "public object f1(ref object a)",
                "{",
                "    object retVal1 = null;",
                "    object byrefalias2 = a;",
                "    try",
                "    {",
                "        _.CALL(_outer, \"f2\", _.ARGS.Ref(byrefalias2, v3 => { byrefalias2 = v3; }));",
                "    }",
                "    finally { a = byrefalias2; }",
                "    return retVal1;",
                "}",
                "public object f2(ref object a)",
                "{",
                "    return null;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a ByRef argument of a function is required within an expression that must potentially trap errors, then an alias of that argument will be
        /// required since the potentially-error-trapped content will be executed within a lambda and C# does not allow "ref" arguments to be accessed
        /// within lambdas. If this alias may be altered - if it is passed into another function as a ByRef argument, for example - then the alias value
        /// must be used to overwrite the original function argument reference, even if the expression evaluation failed (since it might have changed
        /// the value before the error occurred).
        /// </summary>
        [Fact]
        public void ByRefFunctionArgumentMustBeMappedToReadAndWriteAliasIfReferencedInReadAndWriteMannerWithinPotentiallyErrorTrappingStatement()
        {
            var source = @"
                Function F1(a)
                    On Error Resume Next
                    WScript.Echo a
                End Function
            ";
            var expected = new[]
            {
                "public object f1(ref object a)",
                "{",
                "    object retVal1 = null;",
                "    var errOn2 = _.GETERRORTRAPPINGTOKEN();",
                "    _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);",
                "    object byrefalias3 = a;",
                "    try",
                "    {",
                "        _.HANDLEERROR(errOn2, () => {",
                "            _.CALL(_env.wscript, \"echo\", _.ARGS.Ref(byrefalias3, v4 => { byrefalias3 = v4; }));",
                "        });",
                "    }",
                "    finally { a = byrefalias3; }",
                "    _.RELEASEERRORTRAPPINGTOKEN(errOn2);",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a ByRef argument of a function is required within an expression that must potentially trap errors, then an alias of that argument will be
        /// required since the potentially-error-trapped content will be executed within a lambda and C# does not allow "ref" arguments to be accessed
        /// within lambdas. If this alias may not be altered then the alias need not be written back over the original reference, it is a read-only
        /// alias. This would be the case if there is a ByRef argument "a" of the current function and "a.Name" is passed to another function (as a
        /// ByRef OR ByVal argument) since the "a" in "a.Name" can never be affected.
        /// </summary>
        [Fact]
        public void ByRefFunctionArgumentMustBeMappedToReadOnlyAliasIfReferencedInReadOnlyMannerWithinPotentiallyErrorTrappingStatement()
        {
            var source = @"
                Function F1(a)
                    On Error Resume Next
                    WScript.Echo a.Name
                End Function
            ";
            var expected = new[]
            {
                "public object f1(ref object a)",
                "{",
                "    object retVal1 = null;",
                "    var errOn2 = _.GETERRORTRAPPINGTOKEN();",
                "    _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn2);",
                "    object byrefalias3 = a;",
                "    _.HANDLEERROR(errOn2, () => {",
                "        _.CALL(_env.wscript, \"echo\", _.ARGS.Val(_.CALL(byrefalias3, \"name\")));",
                "    });",
                "    _.RELEASEERRORTRAPPINGTOKEN(errOn2);",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
