﻿using CSharpSupport;
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
            var content = @"
                ' Test
                Dim a
                Test1 ' Inline comment
                WScript.Echo 1

                Dim i: For i = 1 To 10
                    WScript.Echo i
                Next";

            // Set to Executable to get an entire, runnable program. Set to WithoutScaffolding to see only the meat of the translated content.
            var outputType = OuterScopeBlockTranslator.OutputTypeOptions.Executable;
            Console.WriteLine(
                Translate(content, outputType)
            );
            Console.ReadLine();
        }

        private static string Translate(string scriptContent, OuterScopeBlockTranslator.OutputTypeOptions outputType)
        {
            if (scriptContent == null)
                throw new ArgumentNullException("scriptContent");
            if ((outputType != OuterScopeBlockTranslator.OutputTypeOptions.Executable) && (outputType != OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding))
                throw new ArgumentOutOfRangeException("outputType");

            // This is just a very simple configuration of the CodeBlockTranslator, its name generation implementations are not robust in
            // the slightest, it's just to get going and should be rewritten when the CodeBlockTranslator is further along functionally
            var random = new Random();
            var startClassName = new CSharpName("TranslatedProgram");
            var startMethodName = new CSharpName("Go");
            var supportRefName = new CSharpName("_");
            var envClassName = new CSharpName("EnvironmentReferences");
            var envRefName = new CSharpName("_env");
            var outerClassName = new CSharpName("GlobalReferences");
            var outerRefName = new CSharpName("_outer");
            VBScriptNameRewriter nameRewriter = name => new CSharpName(DefaultRuntimeSupportClassFactory.DefaultNameRewriter(name.Content));
            TempValueNameGenerator tempNameGenerator = (optionalPrefix, scopeAccessInformation) =>
            {
                return new CSharpName(((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + random.Next(1000000).ToString());
            };
            var logger = new CSharpCommentMakingLogger(
                new ConsoleLogger()
            );
            var statementTranslator = new StatementTranslator(supportRefName, envRefName, outerRefName, nameRewriter, tempNameGenerator, logger);
            var codeBlockTranslator = new OuterScopeBlockTranslator(
                startClassName,
                startMethodName,
                supportRefName,
                envClassName,
                envRefName,
                outerClassName,
                outerRefName,
                nameRewriter,
                tempNameGenerator,
                statementTranslator,
                new ValueSettingStatementsTranslator(supportRefName, envRefName, outerRefName, nameRewriter, statementTranslator, logger),
                new NonNullImmutableList<NameToken>().Add(new NameToken("WScript", 0)),
                outputType,
                logger
            );

            var translatedCodeBlocks = ProcessContent(scriptContent).ToNonNullImmutableList();
            var output = codeBlockTranslator.Translate(translatedCodeBlocks);
            return string.Join(
                "\n",
                output.Select(c => (new string(' ', c.IndentationDepth * 4)) + c.Content)
            );
        }

        /// <summary>
        /// This will wrap log messages in C# comments (ensuring that there is no closing-comment symbol in the content which would invalidate the
        /// output as a comment). If a ConsoleLogger is used and the translated program content is sent to the console then this allows all of the
        /// output to be copy-pasted into a C# file for testing. Pretty rough and ready but can make things a little easier!
        /// </summary>
        private class CSharpCommentMakingLogger : ILogInformation
        {
            private readonly ILogInformation _logger;
            public CSharpCommentMakingLogger(ILogInformation logger)
            {
                if (logger == null)
                    throw new ArgumentNullException("logger");
                _logger = logger;
            }
            public void Warning(string content)
            {
                if (!string.IsNullOrWhiteSpace(content))
                    content = "/* " + content.Replace("*/", "*") + " */";
                _logger.Warning(content);
            }
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
