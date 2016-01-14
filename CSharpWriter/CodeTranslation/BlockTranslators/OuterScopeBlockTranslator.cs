using System;
using System.Collections.Generic;
using System.Linq;
using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Compat;
using CSharpSupport.Exceptions;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
	/// <summary>
	/// The outer scope code blocks need significant rearranging to work with C# compared to how they may be structured in VBScript. In VBScript, any variable declared
	/// in the outer scope (ie. not within a function or a class) may be accessed by any child scope. If there are classes defined and multiple instances of classes,
	/// they can all access references in that outer scope, almost as if those references were static. If there are undeclared variables accessed anywhere then they
	/// act as though they were specifically defined in the outer scope (unless Option Explicit was specified). To avoid having to generate a static class, the approach
	/// taken here is to define an "Outer References" class which has properties for all of the VBScript outer scope variables and contains all of the functions from the
	/// outer scope (since outer scope functions may also be accessed by any VBScript class instances). An instance of this class is created when the generated Start
	/// Method is called and is passed around to any class instances. There is an "Environment References" class which is similar, but for the non-declared variables.
	/// This will not contain any functions. The Start Method will take an argument for this reference so that external objects (eg. Response or WScript) may be
	/// provided, or the parameter-less constructor may be used, in which case a default Environment References instance will be used which leaves all of the values
	/// set to null. This process may require considerable rearranging of the source code since it will pull all of the outer scope's non-function-or-class content
	/// into the Start Method, then declare the Outer and Environment References classes (moving the outer scope functions into the Outer References class) and then
	/// any translated classes. The final Translated Program class will require an IProvideVBScriptCompatFunctionality reference which it will also pass to any
	/// instantiated translated child classes (along with the Outer and Environment references).
	/// </summary>
	public class OuterScopeBlockTranslator : CodeBlockTranslator
	{
		private readonly CSharpName _startClassName, _startMethodName, _runtimeDateLiteralValidatorClassName;
		private readonly NonNullImmutableList<NameToken> _externalDependencies;
		private readonly OutputTypeOptions _outputType;
		private readonly ILogInformation _logger;
		public OuterScopeBlockTranslator(
			CSharpName startClassName,
			CSharpName startMethodName,
			CSharpName runtimeDateLiteralValidatorClassName,
			CSharpName supportRefName,
			CSharpName envClassName,
			CSharpName envRefName,
			CSharpName outerClassName,
			CSharpName outerRefName,
			VBScriptNameRewriter nameRewriter,
			TempValueNameGenerator tempNameGenerator,
			ITranslateIndividualStatements statementTranslator,
			ITranslateValueSettingsStatements valueSettingStatementTranslator,
			NonNullImmutableList<NameToken> externalDependencies,
			OutputTypeOptions outputType,
			ILogInformation logger)
			: base(supportRefName, envClassName, envRefName, outerClassName, outerRefName, nameRewriter, tempNameGenerator, statementTranslator, valueSettingStatementTranslator, logger)
		{
			if (startClassName == null)
				throw new ArgumentNullException("startClassName");
			if (startMethodName == null)
				throw new ArgumentNullException("startMethodName");
			if (runtimeDateLiteralValidatorClassName == null)
				throw new ArgumentNullException("runtimeDateLiteralValidatorClassName");
			if (externalDependencies == null)
				throw new ArgumentNullException("externalDependencies");
			if (!Enum.IsDefined(typeof(OutputTypeOptions), outputType))
				throw new ArgumentOutOfRangeException("outputType");
			if (logger == null)
				throw new ArgumentNullException("logger");

			_startClassName = startClassName;
			_startMethodName = startMethodName;
			_runtimeDateLiteralValidatorClassName = runtimeDateLiteralValidatorClassName;
			_externalDependencies = externalDependencies;
			_outputType = outputType;
			_logger = logger;
		}

		public enum OutputTypeOptions
		{
			/// <summary>
			/// This is the default option, a fully-executable class definition will be returned
			/// </summary>
			Executable,
			
			/// <summary>
			/// This option may be used for testing output since the scaffolding which wraps the translated statement into an executable class is excluded
			/// </summary>
			WithoutScaffolding
		}

		public NonNullImmutableList<TranslatedStatement> Translate(NonNullImmutableList<ICodeBlock> blocks)
		{
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			// There are some date literal values that need to be validated at runtime (they may vary by culture - if they include a month name, basically, which will
			// depend upon the current language). If any of these literals are invalid for the runtime culture then no further processing may take place (this is the
			// same as the VBScript interpreter refusing to process a script when it reads it, in whatever culture is active when it executes). We need to build this
			// list before we remove any duplicate functions since, although VBScript will "overwrite" functions with the same name in the outermost scope, if the
			// functions it overwrites / ignores contained any invalid date literals, the interpreter will still not execute. We'll use this data further down..
			var dateLiteralsToValidateAtRuntime = EnumerateAllDateLiteralTokens(blocks)
				.Where(d => d.RequiresRuntimeValidation)
				.GroupBy(d => d.Content)
				.Select(g => new { DateLiteralValue = g.Key, LineNumbers = g.Select(d => d.LineIndex + 1) })
				.ToArray();

			// Note: There is no need to check for identically-named classes since that would cause a "Name Redefined" error even if Option Explicit was not enabled
			blocks = RemoveDuplicateFunctions(blocks);

			// Group the code blocks that need to be executed (the functions from the outermost scope need to go into a "global references" class
			// which will appear after any other classes, which will appear after everything else)
			var commentBuffer = new NonNullImmutableList<CommentStatement>();
			var annotatedFunctions = new NonNullImmutableList<Annotated<AbstractFunctionBlock>>();
			var annotatedClasses = new NonNullImmutableList<Annotated<ClassBlock>>();
			var otherBlocks = new NonNullImmutableList<ICodeBlock>();
			foreach (var block in blocks)
			{
				var comment = block as CommentStatement;
				if (comment != null)
				{
					commentBuffer = commentBuffer.Add(comment);
					continue;
				}
				var functionBlock = block as AbstractFunctionBlock;
				if (functionBlock != null)
				{
					annotatedFunctions = annotatedFunctions.Add(new Annotated<AbstractFunctionBlock>(commentBuffer, functionBlock));
					commentBuffer = new NonNullImmutableList<CommentStatement>();
					continue;
				}
				var classBlock = block as ClassBlock;
				if (classBlock != null)
				{
					annotatedClasses = annotatedClasses.Add(new Annotated<ClassBlock>(commentBuffer, classBlock));
					commentBuffer = new NonNullImmutableList<CommentStatement>();
					continue;
				}
				otherBlocks = otherBlocks.AddRange(commentBuffer).Add(block);
				commentBuffer = new NonNullImmutableList<CommentStatement>();
			}
			if (commentBuffer.Any())
				otherBlocks = otherBlocks.AddRange(commentBuffer);

			// Ensure that functions are in valid configurations - properties are not valid outside of classes and any non-public functions will
			// be translated INTO public functions (since this there are no private external functions in VBScript)
			annotatedFunctions = annotatedFunctions
				.Select(f =>
				{
					if (f.CodeBlock is PropertyBlock)
						throw new ArgumentException("Property encountered in OuterMostScope - these may only appear within classes: " + f.CodeBlock.Name.Content);
					if (!f.CodeBlock.IsPublic)
					{
						_logger.Warning("OuterScope function \"" + f.CodeBlock.Name.Content + "\" is private, this is invalid and will be changed to public");
						if (f.CodeBlock is FunctionBlock)
						{
							return new Annotated<AbstractFunctionBlock>(
								f.LeadingComments,
								new FunctionBlock(true, f.CodeBlock.IsDefault, f.CodeBlock.Name, f.CodeBlock.Parameters, f.CodeBlock.Statements)
							);
						}
						else if (f.CodeBlock is SubBlock)
						{
							return new Annotated<AbstractFunctionBlock>(
								f.LeadingComments,
								new SubBlock(true, f.CodeBlock.IsDefault, f.CodeBlock.Name, f.CodeBlock.Parameters, f.CodeBlock.Statements)
							);
						}
						else
							throw new ArgumentException("Unsupported AbstractFunctionBlock type: " + f.GetType());
					}
					return f;
				})
				.ToNonNullImmutableList();

			var scopeAccessInformation = ScopeAccessInformation.FromOutermostScope(
				_startClassName, // A placeholder name is required for an OutermostScope instance and so is required by this method
				blocks,
				_externalDependencies
			);
			if (blocks.DoesScopeContainOnErrorResumeNext())
			{
				scopeAccessInformation = scopeAccessInformation.SetErrorRegistrationToken(
					_tempNameGenerator(new CSharpName("errOn"), scopeAccessInformation)
				);
			}

			var outerExecutableBlocksTranslationResult = Translate(
				otherBlocks.ToNonNullImmutableList(),
				scopeAccessInformation,
				3 // indentationDepth
			);
			var explicitVariableDeclarationsFromWithOuterScope = outerExecutableBlocksTranslationResult.ExplicitVariableDeclarations;
			outerExecutableBlocksTranslationResult = new TranslationResult(
				outerExecutableBlocksTranslationResult.TranslatedStatements,
				new NonNullImmutableList<VariableDeclaration>(),
				outerExecutableBlocksTranslationResult.UndeclaredVariablesAccessed
			);

			var translatedStatements = new NonNullImmutableList<TranslatedStatement>();
			if (_outputType == OutputTypeOptions.Executable)
			{
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("using System;", 0, 0),
					new TranslatedStatement("using System.Collections;", 0, 0)
				});
				if (dateLiteralsToValidateAtRuntime.Any())
				{
					// System.Collections.ObjectModel is only required for the ReadOnlyCollection, which is only used when there are date literals that need validating at runtime
					translatedStatements = translatedStatements.Add(new TranslatedStatement("using System.Collections.ObjectModel;", 0, 0));
				}
				translatedStatements = translatedStatements.Add(new TranslatedStatement("using System.Runtime.InteropServices;", 0, 0));
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("using " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Namespace + ";", 0, 0),
					new TranslatedStatement("using " + typeof(SourceClassName).Namespace + ";", 0, 0),
					new TranslatedStatement("using " + typeof(SpecificVBScriptException).Namespace + ";", 0, 0),
					new TranslatedStatement("using " + typeof(TranslatedPropertyIReflectImplementation).Namespace + ";", 0, 0),
					new TranslatedStatement("", 0, 0),
					new TranslatedStatement("namespace " + _startClassName.Name, 0, 0),
					new TranslatedStatement("{", 0, 0),
					new TranslatedStatement("public class Runner", 1, 0),
					new TranslatedStatement("{", 1, 0),
					new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " " + _supportRefName.Name + ";", 2, 0),
					new TranslatedStatement("public Runner(" + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " compatLayer)", 2, 0),
					new TranslatedStatement("{", 2, 0),
					new TranslatedStatement("if (compatLayer == null)", 3, 0),
					new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", 4, 0),
					new TranslatedStatement(_supportRefName.Name + " = compatLayer;", 3, 0),
					new TranslatedStatement("}", 2, 0)
				});
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("", 0, 0),
					new TranslatedStatement(
						string.Format(
							"public {0} {1}()",
							_outerClassName.Name,
							_startMethodName.Name
						),
						2,
						0
					),
					new TranslatedStatement("{", 2, 0),
					new TranslatedStatement(
						string.Format(
							"return {0}(new {1}());",
							_startMethodName.Name,
							_envClassName.Name
						),
						3,
						0
					),
					new TranslatedStatement("}", 2, 0),
					new TranslatedStatement(
						string.Format(
							"public {0} {1}({2} env)",
							_outerClassName.Name,
							_startMethodName.Name,
							_envClassName.Name
						),
						2,
						0
					),
					new TranslatedStatement("{", 2, 0),
					new TranslatedStatement("if (env == null)", 3, 0),
					new TranslatedStatement("throw new ArgumentNullException(\"env\");", 4, 0),
					new TranslatedStatement("", 0, 0),
					new TranslatedStatement(
						string.Format("var {0} = env;", _envRefName.Name),
						3,
						0
					),
					new TranslatedStatement(
						string.Format("var {0} = new {1}({2}, {3});", _outerRefName.Name, _outerClassName.Name, _supportRefName.Name, _envRefName.Name),
						3,
						0
					)
				});
				if (dateLiteralsToValidateAtRuntime.Any())
				{
					// When rendering in full Executable format (not just in WithoutScaffolding), if there were any date literals in the source content that could not
					// be confirmed as valid at translation time (meaning they include a month name, which will vary in validity depending upon the culture of the
					// environment at runtime) then these literals need to be validated each run before any other work is attempted. (This is the equivalent of
					// the VBScript interpreter reading the script for every execution and validating date literals against the current culture - if it finds
					// any of them to be invalid then it will raise a syntax error and not attempt to execute any of the script).
					translatedStatements = translatedStatements.Add(new TranslatedStatement(
						string.Format(
							"{0}.ValidateAgainstCurrentCulture({1});",
							_runtimeDateLiteralValidatorClassName.Name,
							_supportRefName.Name
						),
						3,
						0
					));
				}
				translatedStatements = translatedStatements.Add(new TranslatedStatement("", 0, 0));
			}

			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				translatedStatements = translatedStatements
					.Add(new TranslatedStatement(
						string.Format(
							"var {0} = {1}.GETERRORTRAPPINGTOKEN();",
							scopeAccessInformation.ErrorRegistrationTokenIfAny.Name,
							_supportRefName.Name
						),
						3,
						0
					));
			}
			translatedStatements = translatedStatements.AddRange(
				outerExecutableBlocksTranslationResult.TranslatedStatements
			);
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				translatedStatements = translatedStatements
					.Add(new TranslatedStatement(
						string.Format(
							"{0}.RELEASEERRORTRAPPINGTOKEN({1});",
							_supportRefName.Name,
							scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
						),
						3,
						0
					));
			}
			if (_outputType == OutputTypeOptions.Executable)
			{
				// Close the main "TranslatedProgram" function and then write out the runtime-date-literal-validation logic and the global references class (when a complete executable
				// is not required, none of this is of much benefit and detracts from the core of what's being translated - most tests will specify WithoutScaffolding rather than
				// Executable so that just the real meat of the source code is generated)
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement(string.Format("return {0};", _outerRefName.Name), 3, 0),
					new TranslatedStatement("}", 2, 0),
					new TranslatedStatement("", 0, 0)
				});

				if (dateLiteralsToValidateAtRuntime.Any())
				{
					// Declare a static readonly immutable set of date literals that need validating against the current culture before any request can do any actual work
					translatedStatements = translatedStatements.AddRange(new[]
					{
						new TranslatedStatement(
							string.Format(
								"private static class {0}",
								_runtimeDateLiteralValidatorClassName.Name
							),
							2,
							0
						),
						new TranslatedStatement("{", 2, 0),
						new TranslatedStatement("private static readonly ReadOnlyCollection<Tuple<string, int[]>> _literalsToValidate =", 3, 0),
						new TranslatedStatement("new ReadOnlyCollection<Tuple<string, int[]>>(new[] {", 3, 0)
					});
					foreach (var indexedDateLiteralToValidate in dateLiteralsToValidateAtRuntime.Select((d, i) => new { Index = i, DateLiteralValue = d.DateLiteralValue, LineNumbers = d.LineNumbers }))
					{
						translatedStatements = translatedStatements.Add(new TranslatedStatement(
							string.Format(
								"Tuple.Create({0}, new[] {{ {1} }}){2}",
								indexedDateLiteralToValidate.DateLiteralValue.ToLiteral(),
								string.Join<int>(", ", indexedDateLiteralToValidate.LineNumbers),
								(indexedDateLiteralToValidate.Index < (dateLiteralsToValidateAtRuntime.Length - 1)) ? "," : ""
							),
							4,
							0
						));
					}
					translatedStatements = translatedStatements.Add(new TranslatedStatement("});", 3, 0));
					translatedStatements = translatedStatements.Add(new TranslatedStatement("", 0, 0));

					// Declare the function that reads the data above and performs the validation work
					translatedStatements = translatedStatements.AddRange(new[]
					{
						new TranslatedStatement("public static void ValidateAgainstCurrentCulture(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)", 3, 0),
						new TranslatedStatement("{", 3, 0),
						new TranslatedStatement("if (compatLayer == null)", 4, 0),
						new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", 5, 0),
						new TranslatedStatement("foreach (var dateLiteralValueAndLineNumbers in _literalsToValidate)", 4, 0),
						new TranslatedStatement("{", 4, 0),
						new TranslatedStatement(
							string.Format(
								"try {{ compatLayer.DateLiteralParser.Parse(dateLiteralValueAndLineNumbers.Item1); }}",
								_runtimeDateLiteralValidatorClassName.Name
							),
							5,
							0
						),
						new TranslatedStatement("catch", 5, 0),
						new TranslatedStatement("{", 5, 0),
						new TranslatedStatement("throw new SyntaxError(string.Format(", 6, 0),
						new TranslatedStatement("\"Invalid date literal #{0}# on line{1} {2}\",", 7, 0),
						new TranslatedStatement("dateLiteralValueAndLineNumbers.Item1,", 7, 0),
						new TranslatedStatement("(dateLiteralValueAndLineNumbers.Item2.Length == 1) ? \"\" : \"s\",", 7, 0),
						new TranslatedStatement("string.Join<int>(\", \", dateLiteralValueAndLineNumbers.Item2)", 7, 0),
						new TranslatedStatement("));", 6, 0),
						new TranslatedStatement("}", 5, 0),
						new TranslatedStatement("}", 4, 0),
						new TranslatedStatement("}", 3, 0),
						new TranslatedStatement("}", 2, 0),
						new TranslatedStatement("", 0, 0)
					});
				}

				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("public class " + _outerClassName.Name, 2, 0),
					new TranslatedStatement("{", 2, 0),
					new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " " + _supportRefName.Name + ";", 3, 0),
					new TranslatedStatement("private readonly " + _outerClassName.Name + " " + _outerRefName.Name + ";", 3, 0),
					new TranslatedStatement("private readonly " + _envClassName.Name + " " + _envRefName.Name + ";", 3, 0),
					new TranslatedStatement("public " + _outerClassName.Name + "(" + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " compatLayer, " + _envClassName.Name + " env)", 3, 0),
					new TranslatedStatement("{", 3, 0),
					new TranslatedStatement("if (compatLayer == null)", 4, 0),
					new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", 5, 0),
					new TranslatedStatement("if (env == null)", 4, 0),
					new TranslatedStatement("throw new ArgumentNullException(\"env\");", 5, 0),
					new TranslatedStatement(_supportRefName.Name + " = compatLayer;", 4, 0),
					new TranslatedStatement(_envRefName.Name + " = env;", 4, 0),
					new TranslatedStatement(_outerRefName.Name + " = this;", 4, 0)
				});

				// Note: Any repeated "explicitVariableDeclarationsFromWithOuterScope" entries are ignored - this makes the ReDim translation process easier (where ReDim statements
				// may target already-declared variables or they may be considered to implicitly declare them) but it means that the Dim translation has to do some extra work to
				// pick up on "Name redefined" scenarios.
				var variableInitialisationStatements = new NonNullImmutableList<TranslatedStatement>();
				foreach (var explicitVariableDeclaration in explicitVariableDeclarationsFromWithOuterScope)
				{
					var variableInitialisationStatement = new TranslatedStatement(
						base.TranslateVariableInitialisation(explicitVariableDeclaration, ScopeLocationOptions.OutermostScope),
						4,
						explicitVariableDeclaration.Name.LineIndex
					);
					if (!variableInitialisationStatements.Any(s => s.Content == variableInitialisationStatement.Content))
						variableInitialisationStatements = variableInitialisationStatements.Add(variableInitialisationStatement);
				}
				translatedStatements = translatedStatements.AddRange(variableInitialisationStatements);
				translatedStatements = translatedStatements
					.Add(
						new TranslatedStatement("}", 3, 0)
					);
				if (explicitVariableDeclarationsFromWithOuterScope.Any())
				{
					translatedStatements = translatedStatements.Add(new TranslatedStatement("", 0, 0));
					var variableDeclarationStatements = new NonNullImmutableList<TranslatedStatement>();
					foreach (var explicitVariableDeclaration in explicitVariableDeclarationsFromWithOuterScope)
					{
						var variableDeclarationStatement = new TranslatedStatement(
							"public object " + _nameRewriter.GetMemberAccessTokenName(explicitVariableDeclaration.Name) + " { get; set; }",
							3,
							explicitVariableDeclaration.Name.LineIndex
						);
						if (!variableDeclarationStatements.Any(s => s.Content == variableDeclarationStatement.Content))
							variableDeclarationStatements = variableDeclarationStatements.Add(variableDeclarationStatement);
					}
					translatedStatements = translatedStatements.AddRange(variableDeclarationStatements);
				}
			}
			foreach (var annotatedFunctionBlock in annotatedFunctions)
			{
				// Note that any variables that were accessed by code in the outermost scope, but that were not explicitly declared, are considered to be IMPLICITLY declared.
				// So, if they are referenced within any of the functions, they must share the same reference unless explicit declared within those functions (so if "a" is
				// set in the outermost scope and then a function has a statement "ReDim a(2)", that "a" reference should be that one in the outermost scope). To achieve
				// this, the "implicitly declared" outermost scope variables are added to the ExternalDependencies set in the ScopeAccessInformation instance provided
				// to the function block translator. The same must be done for class block translation (see below).
				translatedStatements = translatedStatements.Add(new TranslatedStatement("", 1, 0));
				translatedStatements = translatedStatements.AddRange(
					Translate(
						annotatedFunctionBlock.LeadingComments.Cast<ICodeBlock>().Concat(new[] { annotatedFunctionBlock.CodeBlock }).ToNonNullImmutableList(),
						scopeAccessInformation.ExtendExternalDependencies(outerExecutableBlocksTranslationResult.UndeclaredVariablesAccessed),
						3 // indentationDepth
					).TranslatedStatements
				);
			}
			var classBlocksTranslationResult = Translate(
				TrimTrailingBlankLines(
					annotatedClasses.SelectMany(c => c.LeadingComments.Cast<ICodeBlock>().Concat(new[] { c.CodeBlock })).ToNonNullImmutableList()
				),
				scopeAccessInformation.ExtendExternalDependencies(outerExecutableBlocksTranslationResult.UndeclaredVariablesAccessed), // See comment above relating to ExternalDependencies for function blocks
				2 // indentationDepth
			);
			if (_outputType == OutputTypeOptions.Executable)
			{
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("}", 2, 0),
					new TranslatedStatement("", 0, 0)
				});

				// This has to be generated after all of the Translate calls to ensure that the UndeclaredVariablesAccessed data for all of the TranslationResults is available
				var allEnvironmentVariablesAccessed =
					scopeAccessInformation.ExternalDependencies
					.AddRange(
						outerExecutableBlocksTranslationResult
							.Add(classBlocksTranslationResult)
							.UndeclaredVariablesAccessed
					);
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("public class " + _envClassName.Name, 2, 0),
					new TranslatedStatement("{", 2, 0)
				});
				var allEnvironmentVariableNames = allEnvironmentVariablesAccessed.Select(v => new { RewrittenName = _nameRewriter.GetMemberAccessTokenName(v), LineIndex = v.LineIndex });
				var environmentVariableNamesThatHaveBeenAccountedFor = new HashSet<string>();
				foreach (var v in allEnvironmentVariableNames)
				{
					if (environmentVariableNamesThatHaveBeenAccountedFor.Contains(v.RewrittenName))
						continue;
					translatedStatements = translatedStatements.Add(
						new TranslatedStatement("public object " + v.RewrittenName + " { get; set; }", 3, v.LineIndex)
					);
					environmentVariableNamesThatHaveBeenAccountedFor.Add(v.RewrittenName);
				}
				translatedStatements = translatedStatements.Add(
					new TranslatedStatement("}", 2, 0)
				);
			}

			if (classBlocksTranslationResult.TranslatedStatements.Any())
			{
				translatedStatements = translatedStatements.Add(
					new TranslatedStatement("", 0, 0)
				);
				translatedStatements = translatedStatements.AddRange(
					classBlocksTranslationResult.TranslatedStatements
				);
			}

			if (_outputType == OutputTypeOptions.Executable)
			{
				translatedStatements = translatedStatements.AddRange(new[]
				{
					new TranslatedStatement("}", 1, 0), // Close outer class
					new TranslatedStatement("}", 0, 0), // Close namespace
				});
			}

			// Moving the functions and classes around can sometimes leaving extraneous blank lines in their wake. This tidies them up. (This would also result in
			// runs of blank lines in the source being reduced to a single line, not just runs that were introduced by the rearranging, but I can't see how that
			// could be a big problem).
			return RemoveRunsOfBlankLines(translatedStatements);
		}

		private NonNullImmutableList<ICodeBlock> TrimTrailingBlankLines(NonNullImmutableList<ICodeBlock> blocks)
		{
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			var result = new NonNullImmutableList<ICodeBlock>();
			foreach (var block in blocks.Reverse())
			{
				var isBlankLine = block is BlankLine;
				if (!isBlankLine || result.Any())
					result = result.Insert(block, 0);
			}
			return result;
		}

		private NonNullImmutableList<TranslatedStatement> RemoveRunsOfBlankLines(NonNullImmutableList<TranslatedStatement> translatedStatements)
		{
			if (translatedStatements == null)
				throw new ArgumentNullException("translatedStatements");

			return translatedStatements
				.Select((s, i) => ((i == 0) || (s.Content != "") || (translatedStatements[i - 1].Content != "")) ? s : null)
				.Where(s => s != null)
				.ToNonNullImmutableList();
		}
		
		private TranslationResult Translate(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			return base.TranslateCommon(
				base.GetWithinFunctionBlockTranslators()
					.Add(base.TryToTranslateClass)
					.Add(base.TryToTranslateFunction)
					.Add(base.TryToTranslateOptionExplicit),
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
		}

		/// <summary>
		/// VBScript allows functions with the same name to appear multiple times, where all but the last implementation will be ignored (this is not
		/// allowed within classes, however properties may exist with the same name as functions and take precedence so long as they come after the
		/// functions - a "Name Redefined" error will be raised if the property comes first or if there are multiple properties with the same name)
		/// </summary>
		private NonNullImmutableList<ICodeBlock> RemoveDuplicateFunctions(NonNullImmutableList<ICodeBlock> blocks)
		{
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			var removeAtLocations = new List<int>();
			foreach (var block in blocks)
			{
				var functionBlock = block as AbstractFunctionBlock;
				if (functionBlock == null)
					continue;

				var functionName = _nameRewriter.GetMemberAccessTokenName(functionBlock.Name);
				removeAtLocations.AddRange(
					blocks
						.Select((b, blockIndex) => new { Index = blockIndex, Block = b })
						.Where(indexedBlock => indexedBlock.Block is AbstractFunctionBlock)
						.Where(indexedBlock => _nameRewriter.GetMemberAccessTokenName(((AbstractFunctionBlock)indexedBlock.Block).Name) == functionName)
						.Select(indexedBlock => indexedBlock.Index)
						.OrderByDescending(blockIndex => blockIndex).Skip(1) // Leave the last one intact
				);
			}
			foreach (var removeIndex in removeAtLocations.Distinct().OrderByDescending(i => i))
				blocks = blocks.RemoveAt(removeIndex);
			return blocks;
		}

		/// <summary>
		/// This will retrieve all DateLiteralToken instances containined in the specified blocks. This may be necessary since there may be some validation
		/// that must be performed before the translated program does any work. There may be date literals, for example, with an English month name in that
		/// will fail when parsed if the program is being executed in a non-English-language environment - in VBScript, this would be a Syntax error at the
		/// point at which the script was parsed, but for programs translated here, we don't want to have to force them to run in the same language as was
		/// used during translation, so date literals are stored in a VBScript-format and checked for validity just before the real execution takes place
		/// (that way, an error may be raised immediately if any of them are no longer valid - emulating the VBScript interpreter's stop-before-executing
		/// behaviour).
		/// </summary>
		private IEnumerable<DateLiteralToken> EnumerateAllDateLiteralTokens(IEnumerable<ICodeBlock> blocks)
		{
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			foreach (var block in blocks)
			{
				if (block == null)
					throw new ArgumentException("Null reference encountered in blocks set");

				IEnumerable<Statement> expressionsToInterrogate;
				var nonNestedExpressionContainingBlock = block as IHaveNonNestedExpressions;
				if (nonNestedExpressionContainingBlock != null)
					expressionsToInterrogate = nonNestedExpressionContainingBlock.NonNestedExpressions;
				else
				{
					var statement = block as Statement;
					if (statement != null)
						expressionsToInterrogate = new[] { statement };
					else
						expressionsToInterrogate = new Statement[0];
				}
				foreach (var dateLiteral in expressionsToInterrogate.SelectMany(e => e.Tokens.OfType<DateLiteralToken>()))
					yield return dateLiteral;

				var nestedContentBlock = block as IHaveNestedContent;
				if (nestedContentBlock != null)
				{
					foreach (var nestedDateLiteral in EnumerateAllDateLiteralTokens(nestedContentBlock.AllExecutableBlocks))
						yield return nestedDateLiteral;
				}
			}
		}

		private class Annotated<T> where T : ICodeBlock
		{
			public Annotated(NonNullImmutableList<CommentStatement> leadingComments, T codeBlock)
			{
				if (leadingComments == null)
					throw new ArgumentNullException("leadingComments");
				if (codeBlock == null)
					throw new ArgumentNullException("codeBlock");

				LeadingComments = leadingComments;
				CodeBlock = codeBlock;
			}

			/// <summary>
			/// This will never be null (though it may be empty)
			/// </summary>
			public NonNullImmutableList<CommentStatement> LeadingComments { get; private set; }

			/// <summary>
			/// This will never be null
			/// </summary>
			public T CodeBlock { get; private set; }
		}
	}
}