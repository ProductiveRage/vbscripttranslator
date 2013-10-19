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
        // =======================================================================================
        // ENTRY POINT
        // =======================================================================================
        static void Main()
        {
            var filename = "Test.vbs";
            var testFileCodeBlocks = ProcessContent(
                GetScriptContent(filename).Replace("\r\n", "\n")
            );
            Console.ReadLine();
        }

        private static IEnumerable<ICodeBlock> ProcessContent(string scriptContent)
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
            Console.WriteLine(getRenderedSourceVB(codeBlocks));
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

        // =======================================================================================
        // READ SCRIPT CONTENT
        // =======================================================================================
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

        // =======================================================================================
        // REBUILD ORIGINAL SOURCE FROM PROCESSED CONTENT
        // =======================================================================================
        private static string getRenderedSourceVB(List<ICodeBlock> content)
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
