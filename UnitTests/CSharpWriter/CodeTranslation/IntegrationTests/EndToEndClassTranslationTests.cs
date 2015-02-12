using System;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndClassTranslationTests
    {
        /// <summary>
        /// When the tokens with the content "Property" had to be classified as a MayBeKeywordOrNameToken instead of a straight KeyWordToken, some logic had to
        /// be changed in the class parsing to account for it - this test exercises that work
        /// </summary>
        [Fact]
        public void EndProperty()
        {
            var source = @"
                CLASS C1
                    PUBLIC PROPERTY GET Name
                    END PROPERTY
                END CLASS
            ";
            var expected = @"
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
                    }

                    public object name()
                    {
                        return null;
                    }
                }";
            Assert.Equal(
                expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a class has a Class_Terminate method with at least one executable statement (ie. not empty and not just blank lines and comments), then it should
        /// implement IDisposable so that it's possible to instantiate it and tidy it up to simulate the way in which the deterministic VBScript interpreter
        /// calls Class_Terminate (as soon as it leaves scope, rather than when a garbage collector wants to deal with it). This isn't currently taken
        /// advantage of in the generated code (as of 2014-12-15) but it might be in the future.
        /// </summary>
        [Fact]
        public void ClassTerminateResultsInDisposableTranslatedClass()
        {
            var source = @"
                CLASS C1
                    PUBLIC SUB Class_Terminate
                        WScript.Echo ""Gone!""
                    END SUB
                END CLASS
            ";
            var expected = @"
                [ComVisible(true)]
                [SourceClassName(""C1"")]
                public sealed class c1 : IDisposable
                {
                    private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
                    private readonly EnvironmentReferences _env;
                    private readonly GlobalReferences _outer;
                    private bool _disposed1;
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
                        _disposed1 = false;
                    }

                    ~c1()
                    {
                        try { Dispose2(false); }
                        catch(Exception e)
                        {
                            try { _env.SETERROR(e); } catch { }
                        }
                    }

                    void IDisposable.Dispose()
                    {
                        Dispose2(true);
                        GC.SuppressFinalize(this);
                    }

                    private void Dispose2(bool disposing)
                    {
                        if (_disposed1)
                            return;
                        if (disposing)
                            class_terminate();
                        _disposed1 = true;
                    }

                    public void class_terminate()
                    {
                        _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Gone!""));
                    }
                }";
            Assert.Equal(
                expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a class has a Class_Initialize method with at least one executable statement (ie. not empty and not just blank lines and comments), then it should
        /// call this method in the constructor in the generated class. For strict compatibility with VBScript, any error is ignored and, while it will terminate
        /// execution of Class_Initialize, it will not prevent the calling code from continuing.
        /// </summary>
        [Fact]
        public void ClassInitializeResultsInConstructorCall()
        {
            var source = @"
                CLASS C1
                    PUBLIC SUB Class_Initialize
                        WScript.Echo ""Here!""
                    END SUB
                END CLASS
            ";
            var expected = @"
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
                        try { class_initialize(); }
                        catch(Exception e)
                        {
                            _env.SETERROR(e);
                        }
                    }

                    public void class_initialize()
                    {
                        _.CALL(_env.wscript, ""echo"", _.ARGS.Val(""Here!""));
                    }
                }";
            Assert.Equal(
                expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
