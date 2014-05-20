using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndReDimTranslationTests
    {
        public class UndeclaredVariables
        {
            [Fact]
            public void NonPreserveReDimOfUndeclaredVariableInTheOutermostScopeShouldImplicitlyDeclareTheVariableInOutermostScope()
            {
                var source = @"
                    ReDim a(0)
                ";
                var expected = new[] {
                    "_outer.a = _.NEWARRAY(0);"
                };
                Assert.Equal(
                    expected,
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void PreserveReDimOfUndeclaredVariableInTheOutermostScopeShouldImplicitlyDeclareTheVariableInOutermostScope()
            {
                var source = @"
                    ReDim Preserve a(0)
                ";
                var expected = new[] {
                    "_outer.a = _.EXTENDARRAY(_outer.a, 0);"
                };
                Assert.Equal(
                    expected,
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void NonPreserveReDimOfUndeclaredVariableInFunctionShouldImplicitlyDeclareTheVariableInLocalScope()
            {
                var source = @"
                    Function F1()
                        ReDim a(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.NEWARRAY(0);
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void PreserveReDimOfUndeclaredVariableInFunctionShouldImplicitlyDeclareTheVariableInLocalScope()
            {
                var source = @"
                    Function F1()
                        ReDim Preserve a(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.EXTENDARRAY(a, 0);
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void NonPreserveReDimOfFunctionReturnValue()
            {
                var source = @"
                    Function F1()
                        ReDim F1(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      retVal1 = _.NEWARRAY(0);
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void PreserveReDimOfFunctionReturnValue()
            {
                var source = @"
                    Function F1()
                        ReDim Preserve F1(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      retVal1 = _.EXTENDARRAY(retVal1, 0);
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }
        }

        public class DeclaredVariables
        {
            [Fact]
            public void NonPreserveReDimOfDeclaredVariableInTheOutermostScope()
            {
                var source = @"
                    Dim a
                    ReDim a(0)
                ";
                var expected = new[] {
                    "_outer.a = _.NEWARRAY(0);"
                };
                Assert.Equal(
                    expected,
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void PreserveReDimOfDeclaredVariableInTheOutermostScope()
            {
                var source = @"
                    Dim a
                    ReDim Preserve a(0)
                ";
                var expected = new[] {
                    "_outer.a = _.EXTENDARRAY(_outer.a, 0);"
                };
                Assert.Equal(
                    expected,
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void NonPreserveReDimOfDeclaredVariableInFunction()
            {
                var source = @"
                    Function F1()
                        Dim a
                        ReDim a(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.NEWARRAY(0);
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            [Fact]
            public void PreserveReDimOfDeclaredVariableInFunction()
            {
                var source = @"
                    Function F1()
                        Dim a
                        ReDim Preserve a(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.EXTENDARRAY(a, 0);
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }
        }

        /// <summary>
        /// ReDim will implicitly declare any target variable, if it has not been already declared - this means that a Dim statement that FOLLOWS a ReDim
        /// will result in a "Name redefined" compile time error in VBScript, so all of these cases should result in a translation exception
        /// </summary>
        public class PrecedingExplicitVariableDeclarations
        {
            [Fact]
            public void NonPreserveReDimOfDeclaredVariableInTheOutermostScope()
            {
                var source = @"
                    ReDim a(0)
                    Dim a
                ";
                Assert.Throws<ArgumentException>(() =>
                {
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
                });
            }

            [Fact]
            public void PreserveReDimOfDeclaredVariableInTheOutermostScope()
            {
                var source = @"
                    ReDim Preserve a(0)
                    Dim a
                ";
                Assert.Throws<ArgumentException>(() =>
                {
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
                });
            }

            [Fact]
            public void NonPreserveReDimOfDeclaredVariableInFunction()
            {
                var source = @"
                    Function F1()
                        ReDim a(0)
                        Dim a
                    End Function";
                Assert.Throws<ArgumentException>(() =>
                {
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
                });
            }

            [Fact]
            public void PreserveReDimOfDeclaredVariableInFunction()
            {
                var source = @"
                    Function F1()
                        ReDim Preserve a(0)
                        Dim a
                    End Function";
                Assert.Throws<ArgumentException>(() =>
                {
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
                });
            }
        }

        private static IEnumerable<string> SplitOnNewLinesSkipFirstLineAndTrimAll(string value)
        {
            if (value == null)
                throw new ArgumentNullException("value");

            return value.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').Skip(1).Select(v => v.Trim());
        }
    }
}
