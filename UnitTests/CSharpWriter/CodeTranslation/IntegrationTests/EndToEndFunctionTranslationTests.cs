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
    }
}
