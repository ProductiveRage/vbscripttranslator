using CSharpWriter.CodeTranslation;
using CSharpWriter.CodeTranslation.BlockTranslators;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
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
            Console.WriteLine(
                Translate(
                    "' Test\n\nDim i\ntest1 ' Inline comment\nWScript.Echo 1"
                )
            );
            Console.ReadLine();
        }

        private static string Translate(string scriptContent)
        {
            if (scriptContent == null)
                throw new ArgumentNullException("scriptContent");

            // This is just a very simple configuration of the CodeBlockTranslator, its name generation implementations are not robust in
            // the slightest, it's just to get going and should be rewritten when the CodeBlockTranslator is further along functionally
            var random = new Random();
            var supportClassName = new CSharpName("_");
            VBScriptNameRewriter nameRewriter = name => new CSharpName(name.Content.ToLower());
            TempValueNameGenerator tempNameGenerator = optionalPrefix => new CSharpName(((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + random.Next(1000000).ToString());
            var logger = new ConsoleLogger();
            var statementTranslator = new StatementTranslator(supportClassName, nameRewriter, tempNameGenerator, logger);
            var codeBlockTranslator = new OuterScopeBlockTranslator(
                supportClassName,
                nameRewriter,
                tempNameGenerator,
                statementTranslator,
                new ValueSettingsStatementsTranslator(supportClassName, nameRewriter, statementTranslator, logger),
                logger
            );

            var translatedCode1 = codeBlockTranslator.Translate(
                ProcessContent(scriptContent).ToNonNullImmutableList()
            );
            return string.Join(
                "\n",
                translatedCode1.Select(c => (new string(' ', c.IndentationDepth * 4)) + c.Content)
            );
        }

        private static IEnumerable<ICodeBlock> ProcessContent(string scriptContent)
        {
            // Translate these tokens into ICodeBlock implementations (representing
            // code VBScript structures)
            string[] endSequenceMet;
            var handler = new CodeBlockHandler(null);
            return handler.Process(
                GetTokens(scriptContent).ToList(),
                out endSequenceMet
            );
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
    }
}
