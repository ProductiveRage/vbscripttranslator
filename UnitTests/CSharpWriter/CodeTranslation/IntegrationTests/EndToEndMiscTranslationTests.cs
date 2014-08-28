﻿using System.Linq;
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
        /// class. There is also a "WScript" reference which is declared as an External Dependency in the translator, this will appear in the Environment
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
    }
}