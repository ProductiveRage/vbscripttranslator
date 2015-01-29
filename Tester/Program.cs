using CSharpWriter;
using CSharpWriter.CodeTranslation.BlockTranslators;
using System;
using System.Linq;

namespace Tester
{
    static class Program
    {
        static void Main()
        {
            // When run, this example code will render an entire class that may be copy-pasted directly into a new file and then executed (so
            // long as the project that it is included in has a reference to "CSharpSupport". The new class has the namer "Runner" and has a
            // constructor that takes a single argument, it requires a "compability functionality provider". An implementation may be found
            // in the static class DefaultRuntimeSupportClassFactory. So executing the translated code may be done with:
            //
            //   using (var compatLayer = DefaultRuntimeSupportClassFactory.Get())
            //   {
            //       new TranslatedProgram.Runner(compatLayer).Go();
            //   }
            //
            // The Go methods takes an optional "EnvironmentReferences" argument. This has properties for all of the undeclared variables in
            // the source code. In the example below, you would need to one with a "WScript" implementation (that has a method "Echo" which
            // takes an object argument and returns a value). Conveniently, such as class is provided in this project, so the output code
            // could be successfully executed with:
            //
            //   using (var compatLayer = DefaultRuntimeSupportClassFactory.Get())
            //   {
            //       new TranslatedProgram.Runner(compatLayer).Go(
            //           new TranslatedProgram.Runner.EnvironmentReferences { WScript = new WScriptMock() }
            //       );
            //   }
            //
            var scriptContent = @"
                ' Test
                Dim i: For i = 1 To 10
                    WScript.Echo ""Item"" & i
                Next";

            // Set to Executable to get an entire, runnable program. Set to WithoutScaffolding to see only the meat of the translated content.
            var outputType = OuterScopeBlockTranslator.OutputTypeOptions.Executable;
            var translatedStatements = DefaultTranslator.Translate(
                scriptContent,
                new[] { "WScript" }, // Assume this is present when translating, don't log warnings about it not being declared
                outputType
            );
            Console.WriteLine(
                string.Join(
                    Environment.NewLine,
                    translatedStatements.Select(c => (new string(' ', c.IndentationDepth * 4)) + c.Content)
                )
            );
            Console.ReadLine();
        }
    }
}