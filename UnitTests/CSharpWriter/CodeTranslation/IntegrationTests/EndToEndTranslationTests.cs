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
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndTranslationTests
    {
        /// <summary>
        /// The code here accesses an undeclared variable in a statement in the outermost scope, that scope should be registered in the EnvironmentReferences
        /// class. There is also a "WScript" reference which is declared as an External Dependency in the translator, this will appear in the Environment
        /// References class as well (as any/all External Dependencies should).
        /// </summary>
        [Fact]
        public void UndeclaredVariablesInTheOutermostScopeShouldBeDefinedAsAnEnvironmentVariable()
        {
            var source = @"
                WScript.Echo i
            ";
            var expected = new[]
            {
                "_.CALL(_env.wscript, \"echo\", new object[] { _env.i });"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                GetTranslatedStatements(source, DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This code will access an undeclared variable within a function. The scope of that undeclared variable should be restricted to the function in
        /// which it is accessed and not bleed out into the outer scope.
        /// </summary>
        [Fact]
        public void UndeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
                Test1
                Function Test1()
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_outer.test1();",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    object i = null; /* Undeclared in source */",
                "    _.CALL(_env.wscript, \"echo\", new object[] { i });",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                GetTranslatedStatements(source, DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [Fact]
        public void DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
                Test1
                Function Test1()
                    Dim i
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_outer.test1();",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    object i = null;",
                "    _.CALL(_env.wscript, \"echo\", new object[] { i });",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                GetTranslatedStatements(source, DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [Fact]
        public void DeclaredVariableInOutermostScopeShouldBeAccessedFromThereWhenRequiredWithinFunction()
        {
            var source = @"
                Dim i
                Test1
                Function Test1()
                    WScript.Echo i
                End Function
            ";
            var expected = new[]
            {
                "_outer.test1();",
                "public object test1()",
                "{",
                "    object retVal1 = null;",
                "    _.CALL(_env.wscript, \"echo\", new object[] { _outer.i });",
                "    return retVal1;",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                GetTranslatedStatements(source, DefaultConsoleExternalDependencies)
            );
        }

        private static NonNullImmutableList<NameToken> DefaultConsoleExternalDependencies
            = new NonNullImmutableList<NameToken>(new[] { new NameToken("WScript", 0) });

        /// <summary>
        /// This will never return null or an array containing any nulls, blank values or values with leading or trailing whitespace or values containing line
        /// returns (this format makes the Assert.Equals easier, where it can make array comparisons easily but not any IEnumerable implementation)
        /// </summary>
        private static string[] GetTranslatedStatements(string content, NonNullImmutableList<NameToken> externalDependencies)
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
            TempValueNameGenerator tempNameGenerator = optionalPrefix => new CSharpName(((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + (++tempNameGeneratorNextNumber).ToString());
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
                new ValueSettingsStatementsTranslator(supportRefName, envRefName, outerRefName, nameRewriter, statementTranslator, logger),
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
