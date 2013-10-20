using CSharpWriter;
using CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;
using VBScriptTranslator.LegacyParser.ContentBreaking;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding;
using VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations;

namespace Tester
{
    static class Program
    {
        static void Main()
        {
            // This is just a very simple configuration of the CodeBlockTranslator, its name generation implementations are not robust in
            // the slightest, it's just to get going and should be rewritten when the CodeBlockTranslator is further along functionally
            var random = new Random();
            var supportClassName = new CSharpName("_");
            VBScriptNameRewriter nameRewriter = name => new CSharpName(name.Content.ToLower());
            TempValueNameGenerator tempNameGenerator = optionalPrefix => new CSharpName(((optionalPrefix == null) ? "" : (optionalPrefix.Name + "_")) + "temp" + random.Next(1000000).ToString());
            var codeBlockTranslator = new CodeBlockTranslator(
                supportClassName,
                nameRewriter,
                tempNameGenerator,
                new StatementTranslator(
                    supportClassName,
                    nameRewriter,
                    tempNameGenerator
                )
            );

            var translatedCode1 = codeBlockTranslator.Translate(
                ProcessContent(
                    "' Test\n\nDim i\ntest1 ' Inline comment\nWScript.Echo 1",
                    false
                ).ToNonNullImmutableList()
            );
            Console.WriteLine(
                string.Join(
                    "\n",
                    translatedCode1.Select(c => (new string(' ', c.IndentationDepth * 4)) + c.Content)
                )
            );
            Console.ReadLine();

            var filename = "Test.vbs";
            var testFileCodeBlocks = ProcessContent(
                GetScriptContent(filename).Replace("\r\n", "\n"),
                true
            );
            Console.ReadLine();

            var translatedCode2 = codeBlockTranslator.Translate(
                testFileCodeBlocks.ToNonNullImmutableList()
            );
            Console.ReadLine();
        }

        private static IEnumerable<ICodeBlock> ProcessContent(string scriptContent, bool pushContentToConsole)
        {
            // Translate these tokens into ICodeBlock implementations (representing
            // code VBScript structures)
            string[] endSequenceMet;
            var handler = new CodeBlockHandler(null);
            var codeBlocks = handler.Process(
                GetTokens(scriptContent).ToList(),
                out endSequenceMet
            );

            // DEBUG: ender processed content as VBScript source code
            if (pushContentToConsole)
                Console.WriteLine(GetRenderedSourceVB(codeBlocks));
            return codeBlocks;
        }

        private static IEnumerable<IToken> GetTokens(string scriptContent)
        {
            // Break down content into String, Comment and UnprocessedContent tokens
            var tokens = StringBreaker.SegmentString(scriptContent);

            // Break down further into String, Comment, Atom and AbstractEndOfStatement tokens
            var atomTokens = new List<IToken>();
            foreach (var token in tokens)
            {
                if (token is UnprocessedContentToken)
                    atomTokens.AddRange(TokenBreaker.BreakUnprocessedToken((UnprocessedContentToken)token));
                else
                    atomTokens.Add(token);
            }

            return NumberRebuilder.Rebuild(OperatorCombiner.Combine(atomTokens)).ToList();
        }

        private static string GetScriptContent(string filename)
        {
            if ((filename ?? "").Trim() == "")
                throw new ArgumentException("GetScriptContent: filename is null or blank");
            
            var file = new FileInfo(filename);
            using (var stream = file.OpenText())
            {
                return stream.ReadToEnd();
            }
        }

        private static string GetRenderedSourceVB(List<ICodeBlock> content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            
            var output = new StringBuilder();
            for (var index = 0; index < content.Count; index++)
            {
                var block = content[index];
                var blockNext = (content.Count > (index + 1)) ? content[index + 1] : null;
                output.Append(block.GenerateBaseSource(new SourceIndentHandler()));
                if (blockNext is InlineCommentStatement)
                {
                    output.Append(" " + blockNext.GenerateBaseSource(new NullIndenter()));
                    index++;
                }
                output.AppendLine();
            }
            return output.ToString();
        }
    }
}
