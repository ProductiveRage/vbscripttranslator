using CSharpSupport;
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

namespace CSharpWriter
{
    public class DefaultTranslator
    {
        /// <summary>
        /// This will attempt to translate VBScript content into C# using the default configurations, probably the best place to start (it uses the
        /// DefaultRuntimeSupportClassFactory for name rewriting, so that same name rewriter must be used to execute the output generated here). If
        /// there are any runtime references that are known to be present (such as WScript when run within CScript at the command line, or Request,
        /// Response, Session, etc.. when run within ASP) then specify their names in the externalDependencies set - this will prevent warnings
        /// being logged in relation to the absence of their definition in the source.
        /// </summary>
        public static NonNullImmutableList<TranslatedStatement> Translate(
            string scriptContent,
            NonNullImmutableList<string> externalDependencies,
            OuterScopeBlockTranslator.OutputTypeOptions outputType,
            bool renderCommentsAboutUndeclaredVariables = true)
        {
            if (scriptContent == null)
                throw new ArgumentNullException("scriptContent");
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");
            if ((outputType != OuterScopeBlockTranslator.OutputTypeOptions.Executable) && (outputType != OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding))
                throw new ArgumentOutOfRangeException("outputType");

            var startClassName = new CSharpName("TranslatedProgram");
            var startMethodName = new CSharpName("Go");
            var supportRefName = new CSharpName("_");
            var envClassName = new CSharpName("EnvironmentReferences");
            var envRefName = new CSharpName("_env");
            var outerClassName = new CSharpName("GlobalReferences");
            var outerRefName = new CSharpName("_outer");
            VBScriptNameRewriter nameRewriter = name => new CSharpName(DefaultRuntimeSupportClassFactory.DefaultNameRewriter(name.Content));
            var tempNameGeneratorNextNumber = 0;
            TempValueNameGenerator tempNameGenerator = (optionalPrefix, scopeAccessInformation) =>
            {
                // To get unique names for any given translation, a running counter is maintained and appended to the end of the generated
                // name. This is only run during translation (this code is not used during execution) so there will be a finite number of
                // times that this is called (so there should be no need to worry about the int value overflowing!)
                return new CSharpName(((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + (++tempNameGeneratorNextNumber).ToString());
            };
            ILogInformation logger;
            if (renderCommentsAboutUndeclaredVariables)
                logger = new CSharpCommentMakingLogger(new ConsoleLogger());
            else
                logger = new NullLogger();
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
                externalDependencies.Select(name => new NameToken(name, 0)).ToNonNullImmutableList(),
                outputType,
                logger
            );

            return codeBlockTranslator.Translate(
                ProcessContent(scriptContent).ToNonNullImmutableList()
            );
        }

        // This alternate signature is just to make the very outer layer a bit friendlier (since it will be used in examples, that seems worthwhile)
        public static NonNullImmutableList<TranslatedStatement> Translate(
            string scriptContent,
            string[] externalDependencies,
            OuterScopeBlockTranslator.OutputTypeOptions outputType,
            bool renderCommentsAboutUndeclaredVariables = true)
        {
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");

            return Translate(scriptContent, externalDependencies.ToNonNullImmutableList(), outputType);
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
