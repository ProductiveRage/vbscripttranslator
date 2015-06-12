using System;
using System.Collections.Generic;
using System.Linq;
using CSharpWriter;
using CSharpWriter.CodeTranslation.BlockTranslators;
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
                    "_.NEWARRAY(new object[] { (Int16)0 }, value1 => { _outer.a = value1; });"
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
                    "_.RESIZEARRAY(_outer.a, new object[] { (Int16)0 }, value1 => { _outer.a = value1; });"
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
                      _.NEWARRAY(new object[] { (Int16)0 }, value2 => { a = value2; });
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
                      _.RESIZEARRAY(a, new object[] { (Int16)0 }, value2 => { a = value2; });
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
                      _.NEWARRAY(new object[] { (Int16)0 }, value2 => { retVal1 = value2; });
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
                      _.RESIZEARRAY(retVal1, new object[] { (Int16)0 }, value2 => { retVal1 = value2; });
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// This test is just to ensure that multiple ReDim statements for the same otherwise-undeclared variable do not result in that variable
            /// being defined multiple times in the C# code (when the ReDim statements exist within in the outermost scope)
            /// </summary>
            [Fact]
            public void RepeatedReDimInOutermostScope()
            {
                var source = @"
                    ReDim a(0)
                    ReDim a(1)
                    ReDim a(2)";
                
                var trimmedTranslatedStatements = DefaultTranslator.Translate(source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable)
                    .Select(s => s.Content.Trim())
                    .ToArray();

                Assert.Equal(1, trimmedTranslatedStatements.Count(s => s == "a = null;"));
                Assert.Equal(1, trimmedTranslatedStatements.Count(s => s == "public object a { get; set; }"));
            }

            /// <summary>
            /// This test is just to ensure that multiple ReDim statements for the same otherwise-undeclared variable do not result in that variable
            /// being defined multiple times in the C# code (when the ReDim statements exist within a function or property)
            /// </summary>
            [Fact]
            public void RepeatedReDimInFunction()
            {
                var source = @"
                    Function F1()
                        ReDim a(0)
                        ReDim a(1)
                        ReDim a(2)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      _.NEWARRAY(new object[] { (Int16)0 }, value2 => { a = value2; });
                      _.NEWARRAY(new object[] { (Int16)1 }, value3 => { a = value3; });
                      _.NEWARRAY(new object[] { (Int16)2 }, value4 => { a = value4; });
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
                    "_.NEWARRAY(new object[] { (Int16)0 }, value1 => { _outer.a = value1; });"
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
                    "_.RESIZEARRAY(_outer.a, new object[] { (Int16)0 }, value1 => { _outer.a = value1; });"
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
                      _.NEWARRAY(new object[] { (Int16)0 }, value2 => { a = value2; });
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
                      _.RESIZEARRAY(a, new object[] { (Int16)0 }, value2 => { a = value2; });
                      return retVal1;
                    }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// This is almost identical to the corresponding test in the UndeclaredVariables class but it ensure that a Dim statement before the repeated
            /// ReDims does not cause any problems (or, in fact, change in behaviour)
            /// </summary>
            [Fact]
            public void RepeatedReDimInOutermostScope()
            {
                var source = @"
                    Dim a
                    ReDim a(0)
                    ReDim a(1)
                    ReDim a(2)";

                var trimmedTranslatedStatements = DefaultTranslator.Translate(source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable)
                    .Select(s => s.Content.Trim())
                    .ToArray();

                Assert.Equal(1, trimmedTranslatedStatements.Count(s => s == "a = null;"));
                Assert.Equal(1, trimmedTranslatedStatements.Count(s => s == "public object a { get; set; }"));
            }

            /// <summary>
            /// This is almost identical to the corresponding test in the UndeclaredVariables class but it ensure that a Dim statement before the repeated
            /// ReDims does not cause any problems (or, in fact, change in behaviour)
            /// </summary>
            [Fact]
            public void RepeatedReDimInFunction()
            {
                var source = @"
                    Function F1()
                        Dim a
                        ReDim a(0)
                        ReDim a(1)
                        ReDim a(2)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      _.NEWARRAY(new object[] { (Int16)0 }, value2 => { a = value2; });
                      _.NEWARRAY(new object[] { (Int16)1 }, value3 => { a = value3; });
                      _.NEWARRAY(new object[] { (Int16)2 }, value4 => { a = value4; });
                      return retVal1;
                   }";
                Assert.Equal(
                    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// A "Dim a()" will result in an explicit array-type variable declaration while a subsequent "ReDim a(0)" will result in an explicit non-array-type
            /// variable declaration (followed by an array initialisation targetting that variable). The non-array-type variable declaration from the ReDim must
            /// be ignored, the array-type declaration from the Dim must take precedence.
            /// </summary>
            [Fact]
            public void ReDimFollowingNonDimensionalArrayDimInFunction()
            {
                var source = @"
                    Function F1()
                        Dim a()
                        ReDim a(0)
                    End Function";
                var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = (object[])null;
                      _.NEWARRAY(new object[] { (Int16)0 }, value2 => { a = value2; });
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

            /// <summary>
            /// If a ReDim exists for a particular variable before a Dim for the same variable, even if they are not present on a single code branch that may
            /// be executed by a single request, the Dim will still result in a "Name redefined" error being raise
            /// </summary>
            [Fact]
            public void ReDimBeforeDimButOnDifferentCodePath()
            {
                var source = @"
                    Function F1()
                        If (True) Then
                            ReDim a(0)
                        Else
                            Dim a
                        End If
                    End Function";
                Assert.Throws<ArgumentException>(() =>
                {
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
                });
            }
        }

        /// <summary>
        /// While a REDIM statement may be interpreted as explicitly declaring a variable when its target variable has not been declared already in any accessible scope, if there IS
        /// a variable that it might be referencing in a parent scope then the REDIM should NOT be interpreted as explicitly declaring a new variable (even if the variable in the
        /// parent scope was only IMPLICITLY declared - ie. accessed but never DIM'd)
        /// </summary>
        [Fact]
        public void ReDimsWithinFunctionCanPointToImplicitlyDeclaredOuterMostScopeVariables()
        {
            var source = @"
                a = 1
                Function F1()
                    ReDim a(2) ' This refers to the implicitly-declared variable ""a"" in the outermost scope
                End Function
                Class C1
                    Private c
                    Function CF1()
                        ReDim a(3) ' This refers to the implicitly-declared variable ""a"" in the outermost scope
                        ReDim b(3) ' There is no reference for this to relate to, so it acts as new explicit variable declaration
                        ReDim c(3) ' This refers to the explicitly-declared variable ""c"" in the containing class
                    End Function
                End Class";
            var expected = @"
                _env.a = (Int16)1;
                public object f1()
                {
                    object retVal1 = null;
                    _.NEWARRAY(new object[] { (Int16)2 }, value2 => { _env.a = value2; }); // This refers to the implicitly-declared variable ""a"" in the outermost scope
                    return retVal1;
                }
                [ComVisible(true)]
                [SourceClassName(""C1"")]
                public sealed class c1
                {
                    private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
                    private readonly EnvironmentReferences _env;
                    private readonly GlobalReferences _outer;
                    public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
                    {
                        if (compatLayer == null)
                            throw new ArgumentNullException(""compatLayer"");
                        if (env == null)
                            throw new ArgumentNullException(""env"");
                        if (outer == null)
                            throw new ArgumentNullException(""outer"");
                        _ = compatLayer;
                        _env = env;
                        _outer = outer;
                        c = null;
                    }
                    private object c { get; set; }
                    public object cf1()
                    {
                        object retVal3 = null;
                        object b = null;
                        _.NEWARRAY(new object[] { (Int16)3 }, value4 => { _env.a = value4; }); // This refers to the implicitly-declared variable ""a"" in the outermost scope
                        _.NEWARRAY(new object[] { (Int16)3 }, value5 => { b = value5; }); // There is no reference for this to relate to, so it acts as new explicit variable declaration
                        _.NEWARRAY(new object[] { (Int16)3 }, value6 => { c = value6; }); // This refers to the explicitly-declared variable ""c"" in the containing class
                        return retVal3;
                    }
                }";
            Assert.Equal(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        private static IEnumerable<string> SplitOnNewLinesSkipFirstLineAndTrimAll(string value)
        {
            if (value == null)
                throw new ArgumentNullException("value");

            return value.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').Skip(1).Select(v => v.Trim());
        }
    }
}
