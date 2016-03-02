using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.CSharpWriter.CodeTranslation;
using VBScriptTranslator.CSharpWriter.CodeTranslation.BlockTranslators;
using VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation;
using VBScriptTranslator.CSharpWriter.Lists;
using VBScriptTranslator.CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.ContentBreaking;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding;
using VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations;

namespace VBScriptTranslator.CSharpWriter
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
			ILogInformation logger;
			if (renderCommentsAboutUndeclaredVariables)
				logger = new CSharpCommentMakingLogger(new ConsoleLogger());
			else
				logger = new NullLogger();

			return Translate(scriptContent, externalDependencies, outputType, logger);
		}

		/// <summary>
		/// This Translate signature exists to provide an extremely simple way to get code translated - it is used in some of the examples so that
		/// there's a way to get to translating before worrying about what the NonNullImmutableList type is all about
		/// </summary>
		public static NonNullImmutableList<TranslatedStatement> Translate(
			string scriptContent,
			string[] externalDependencies,
			OuterScopeBlockTranslator.OutputTypeOptions outputType = OuterScopeBlockTranslator.OutputTypeOptions.Executable,
			bool renderCommentsAboutUndeclaredVariables = true)
		{
			if (externalDependencies == null)
				throw new ArgumentNullException("externalDependencies");

			return Translate(scriptContent, externalDependencies.ToNonNullImmutableList(), outputType);
		}

		/// <summary>
		/// This Translate signature exists to provide a slightly-simpler way to specify a custom warning logger (by providing a simple delegate,
		/// rather than having to provide an ILogInformation implementation)
		/// </summary>
		public static NonNullImmutableList<TranslatedStatement> Translate(
			string scriptContent,
			string[] externalDependencies,
			Action<string> warningLogger,
			OuterScopeBlockTranslator.OutputTypeOptions outputType = OuterScopeBlockTranslator.OutputTypeOptions.Executable)
		{
			if (externalDependencies == null)
				throw new ArgumentNullException(nameof(externalDependencies));
			if (warningLogger == null)
				throw new ArgumentNullException("warningLogger");

			return Translate(scriptContent, externalDependencies.ToNonNullImmutableList(), outputType, new DelegateWrappingWarningLogger(warningLogger));
		}

		/// <summary>
		/// This Translate signature is what the others call into - it doesn't try to hide the fact that externalDependencies should be a NonNullImmutableList
		/// of strings and it requires an ILogInformation implementation to deal with logging warnings
		/// </summary>
		public static NonNullImmutableList<TranslatedStatement> Translate(
			string scriptContent,
			NonNullImmutableList<string> externalDependencies,
			OuterScopeBlockTranslator.OutputTypeOptions outputType,
			ILogInformation logger)
		{
			if (scriptContent == null)
				throw new ArgumentNullException("scriptContent");
			if (externalDependencies == null)
				throw new ArgumentNullException("externalDependencies");
			if ((outputType != OuterScopeBlockTranslator.OutputTypeOptions.Executable) && (outputType != OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding))
				throw new ArgumentOutOfRangeException("outputType");
			if (logger == null)
				throw new ArgumentNullException(nameof(logger));

			var startNamespace = new CSharpName("TranslatedProgram");
			var startClassName = new CSharpName("Runner");
			var startMethodName = new CSharpName("Go");
			var runtimeDateLiteralValidatorClassName = new CSharpName("RuntimeDateLiteralValidator");
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
			var statementTranslator = new StatementTranslator(supportRefName, envRefName, outerRefName, nameRewriter, tempNameGenerator, logger);
			var codeBlockTranslator = new OuterScopeBlockTranslator(
				startNamespace,
				startClassName,
				startMethodName,
				runtimeDateLiteralValidatorClassName,
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
				Parse(scriptContent).ToNonNullImmutableList()
			);
		}
		
		/// <summary>
		/// This will return just the parsed VBScript content, it will not attempt any translation. It will never return null nor a set containing
		/// any null references. This may be used to analyse the structure of a script, if so desired.
		/// </summary>
		public static IEnumerable<ICodeBlock> Parse(string scriptContent)
		{
			// Translate these tokens into ICodeBlock implementations (representing code VBScript structures)
			string[] endSequenceMet;
			var handler = new CodeBlockHandler(null);
			return handler.Process(
				GetTokens(scriptContent).ToList(),
				out endSequenceMet
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

		private sealed class DelegateWrappingWarningLogger : ILogInformation
		{
			private readonly Action<string> _warningLogger;
			public DelegateWrappingWarningLogger(Action<string> warningLogger)
			{
				if (warningLogger == null)
					throw new ArgumentNullException(nameof(warningLogger));

				_warningLogger = warningLogger;
			}

			public void Warning(string content)
			{
				_warningLogger(content);
			}
		}
	}
}
