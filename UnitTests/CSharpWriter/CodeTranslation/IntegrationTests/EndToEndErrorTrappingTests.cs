using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndErrorTrappingTests
    {
        /// <summary>
        /// This is the most basic example - a single OnErrorResumeNext that applies to a single statement that follows it. Whenever any scope terminates,
        /// any error token must be released, which 
        /// </summary>
        [Fact]
        public void SingleErrorTrappedStatement()
        {
            var source = @"
                On Error Resume Next
                WScript.Echo ""Test1""
            ";
            var expected = @"
                var errOn1 = _.GETERRORTRAPPINGTOKEN();
                _.STARTERRORTRAPPING(errOn1);
                _.HANDLEERROR(errOn1, () => {
                    _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test1""));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn1);";
            Assert.Equal(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If an error token is required, it will also be defined at the top of the scope, not just before the first OnErrorResumeNext (in case it is
        /// required elsewhere in the same VBScript scope but in a different C# block scope in the translated output)
        /// </summary>
        [Fact]
        public void FlatStatementSetWithMiddleOneErrorTrapped()
        {
            var source = @"
                WScript.Echo ""Test1""
                On Error Resume Next
                WScript.Echo ""Test2""
                On Error Goto 0
                WScript.Echo ""Test3""
            ";
            var expected = @"
                var errOn1 = _.GETERRORTRAPPINGTOKEN();
                _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test1""));
                _.STARTERRORTRAPPING(errOn1);
                _.HANDLEERROR(errOn1, () => {
                    _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test2""));
                });
                _.STOPERRORTRAPPING(errOn1);
                _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test3""));
                _.RELEASEERRORTRAPPINGTOKEN(errOn1);";
            Assert.Equal(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// Although the condition around the OnErrorResumeNext can never be met, the following statement will still have the error-trapping code around
        /// it since the analysis of the code paths only checks for what look like the potential to enable error-trapping. Since the condition is always
        /// false, in the translated code the STARTERRORTRAPPING call will not be made and so the HANDLEERROR will perform no work (it will not trap the
        /// error) but it is a layer of redirection that can be avoided if the translator is sure that it's not required (if there was no OnErrorResumeNext
        /// present all, for example).
        /// </summary>
        [Fact]
        public void ErrorTrappingLayerMustBeAddedEvenIfItWillOnlyPotentiallyBeEnabled()
        {
            var source = @"
                If (False) Then
                    On Error Resume Next
                End If
                WScript.Echo ""Test1""
            ";
            var expected = @"
                var errOn1 = _.GETERRORTRAPPINGTOKEN();
                if (_.IF(_.Constants.False))
                {
                    _.STARTERRORTRAPPING(errOn1);
                }
                _.HANDLEERROR(errOn1, () => {
                    _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test1""));
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn1);";
            Assert.Equal(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void ErrorTrappingDoesNotAffectChildScopes()
        {
            var source = @"
                On Error Resume Next
                Func1
                Function Func1()
                    WScript.Echo ""Test1""
                End Function
            ";
            var expected = @"
                var errOn1 = _.GETERRORTRAPPINGTOKEN();
                _.STARTERRORTRAPPING(errOn1);
                _.HANDLEERROR(errOn1, () => {
                    _.CALL(_outer, ""func1"");
                });
                _.RELEASEERRORTRAPPINGTOKEN(errOn1);
                public object func1()
                {
                    object retVal2 = null;
                    _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test1""));
                    return retVal2;
                }";
            Assert.Equal(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void ErrorTrappingDoesNotAffectParentScopes()
        {
            var source = @"
                Func1
                WScript.Echo ""Test2""
                Function Func1()
                    On Error Resume Next
                    WScript.Echo ""Test1""
                End Function
            ";
            var expected = @"
                _.CALL(_outer, ""func1"");
                _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test2""));
                public object func1()
                {
                    object retVal1 = null;
                    var errOn2 = _.GETERRORTRAPPINGTOKEN();
                    _.STARTERRORTRAPPING(errOn2);
                    _.HANDLEERROR(errOn2, () => {
                        _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Test1""));
                    });
                    _.RELEASEERRORTRAPPINGTOKEN(errOn2);
                    return retVal1;
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
