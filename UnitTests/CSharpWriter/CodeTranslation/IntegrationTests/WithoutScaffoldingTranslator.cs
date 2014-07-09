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

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public static class WithoutScaffoldingTranslator
    {
        public static NonNullImmutableList<NameToken> DefaultConsoleExternalDependencies
            = new NonNullImmutableList<NameToken>(new[] { new NameToken("WScript", 0) });

        /// <summary>
        /// This will never return null or an array containing any nulls, blank values or values with leading or trailing whitespace or values containing line
        /// returns (this format makes the Assert.Equals easier, where it can make array comparisons easily but not any IEnumerable implementation)
        /// </summary>
        public static string[] GetTranslatedStatements(string content, NonNullImmutableList<NameToken> externalDependencies)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");

            return GetTranslator(externalDependencies)
                .Translate(
                    GetCodeBlocksFromScript(content)
                )
                .Select(s => s.Content)
                .Where(s => s != "")
                .ToArray();
        }

        private static OuterScopeBlockTranslator GetTranslator(NonNullImmutableList<NameToken> externalDependencies)
        {
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");

            var startClassName = new CSharpName("TranslatedProgram");
            var startMethodName = new CSharpName("Go");
            var supportRefName = new CSharpName("_");
            var envClassName = new CSharpName("EnvironmentReferences");
            var envRefName = new CSharpName("_env");
            var outerClassName = new CSharpName("GlobalReferences");
            var outerRefName = new CSharpName("_outer");
            VBScriptNameRewriter nameRewriter = name => new CSharpName(name.Content.ToLower());
            var tempNameGeneratorNextNumber = 0;
            TempValueNameGenerator tempNameGenerator = (optionalPrefix, scopeAccessInformation) =>
            {
                return new CSharpName(((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + (++tempNameGeneratorNextNumber).ToString());
            };
            var logger = new NullLogger();
            var statementTranslator = new StatementTranslator(supportRefName, envRefName, outerRefName, nameRewriter, tempNameGenerator, logger);
            return new OuterScopeBlockTranslator(
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
                externalDependencies,
                OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding,
                logger
            );
        }

        private static NonNullImmutableList<ICodeBlock> GetCodeBlocksFromScript(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            // Translate these tokens into ICodeBlock implementations (representing code VBScript structures)
            string[] endSequenceMet;
            return new CodeBlockHandler(null)
                .Process(
                    GetTokens(content).ToList(),
                    out endSequenceMet
                )
                .ToNonNullImmutableList();
        }

        private static IEnumerable<IToken> GetTokens(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            // Break down content into String, Comment and UnprocessedContent tokens
            var tokens = StringBreaker.SegmentString(content);

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
