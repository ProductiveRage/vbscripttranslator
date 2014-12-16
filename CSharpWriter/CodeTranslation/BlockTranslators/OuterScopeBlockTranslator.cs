using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
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
        private readonly CSharpName _startClassName, _startMethodName;
        private readonly NonNullImmutableList<NameToken> _externalDependencies;
        private readonly OutputTypeOptions _outputType;
        private readonly ILogInformation _logger;
        public OuterScopeBlockTranslator(
            CSharpName startClassName,
            CSharpName startMethodName,
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
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");
            if (!Enum.IsDefined(typeof(OutputTypeOptions), outputType))
                throw new ArgumentOutOfRangeException("outputType");
            if (logger == null)
                throw new ArgumentNullException("logger");

            _startClassName = startClassName;
            _startMethodName = startMethodName;
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

            // Note: There is no need to check for identically-named classes since that would cause a "Name Redefined" error even if Option Explicit was not enabled
            blocks = RemoveDuplicateFunctions(blocks);

            // Group the code blocks that need to be executed;
            Func<ICodeBlock, bool> isClassBlock = block => block is ClassBlock;
            Func<ICodeBlock, bool> isFunctionBlock = block => block is AbstractFunctionBlock;
            var classBlocks = blocks.Where(isClassBlock);
            var functionBlocks = blocks.Where(isFunctionBlock).Cast<AbstractFunctionBlock>().Select(f =>
            {
                // Ensure that functions are in valid configurations - properties are not valid outside of classes and any non-public functions will
                // be translated INTO public functions (since this there are no private external functions in VBScript)
                if (f is PropertyBlock)
                    throw new ArgumentException("Property encountered in OuterMostScope - these may only appear within classes: " + f.Name.Content);
                if (!f.IsPublic)
                {
                    _logger.Warning("OuterScope function \"" + f.Name.Content + "\" is private, this is invalid and will be changed to public");
                    if (f is FunctionBlock)
                        return new FunctionBlock(true, f.IsDefault, f.Name, f.Parameters, f.Statements);
                    else if (f is SubBlock)
                        return new SubBlock(true, f.IsDefault, f.Name, f.Parameters, f.Statements);
                    else
                        throw new ArgumentException("Unsupported AbstractFunctionBlock type: " + f.GetType());
                }
                return f;
            });
            var other = blocks.Where(block => !isClassBlock(block) && !isFunctionBlock(block));

            // TODO: The function and class (and any other) rearranging could be a problem with comments, try to do something about that?
            // - eg. if there are comments for a function on the lines just before the function (outside the function rather than inside it) then they
            //   will get left behind when the functions are moved down
            // - or just remove comments entirely since the translated code is bearing less and less relation to the source?

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
                other.ToNonNullImmutableList(),
                scopeAccessInformation,
                3 // indentationDepth
            );
            var explicitVariableDeclarationsFromWithOuterScope = outerExecutableBlocksTranslationResult.ExplicitVariableDeclarations;
            base.ThrowExceptionForDuplicateVariableDeclarationNames(explicitVariableDeclarationsFromWithOuterScope);
            outerExecutableBlocksTranslationResult = new TranslationResult(
                outerExecutableBlocksTranslationResult.TranslatedStatements,
                new NonNullImmutableList<VariableDeclaration>(),
                outerExecutableBlocksTranslationResult.UndeclaredVariablesAccessed
            );
            var classBlocksTranslationResult = Translate(
                TrimTrailingBlankLines(classBlocks.ToNonNullImmutableList()),
                scopeAccessInformation,
                2 // indentationDepth
            );

            var translatedStatements = new NonNullImmutableList<TranslatedStatement>();
            if (_outputType == OutputTypeOptions.Executable)
            {
                translatedStatements = translatedStatements.AddRange(new[]
                {
                    new TranslatedStatement("using System;", 0),
                    new TranslatedStatement("using System.Runtime.InteropServices;", 0),
                    new TranslatedStatement("using " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Namespace + ";", 0),
                    new TranslatedStatement("using " + typeof(SourceClassName).Namespace + ";", 0),
                    new TranslatedStatement("", 0),
                    new TranslatedStatement("namespace " + _startClassName.Name, 0),
                    new TranslatedStatement("{", 0),
                    new TranslatedStatement("public class Runner", 1),
                    new TranslatedStatement("{", 1),
                    new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " " + _supportRefName.Name + ";", 2),
                    new TranslatedStatement("public Runner(" + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " compatLayer)", 2),
                    new TranslatedStatement("{", 2),
                    new TranslatedStatement("if (compatLayer == null)", 3),
                    new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", 4),
                    new TranslatedStatement(_supportRefName.Name + " = compatLayer;", 3),
                    new TranslatedStatement("}", 2),
                    new TranslatedStatement("", 0),
                    new TranslatedStatement(
                        string.Format(
                            "public void {0}()",
                            _startMethodName.Name
                        ),
                        2
                    ),
                    new TranslatedStatement("{", 2),
                    new TranslatedStatement(
                        string.Format(
                            "{0}(new {1}());",
                            _startMethodName.Name,
                            _envClassName.Name
                        ),
                        3
                    ),
                    new TranslatedStatement("}", 2),
                    new TranslatedStatement(
                        string.Format(
                            "public void {0}({1} env)",
                            _startMethodName.Name,
                            _envClassName.Name
                        ),
                        2
                    ),
                    new TranslatedStatement("{", 2),
                    new TranslatedStatement("if (env == null)", 3),
                    new TranslatedStatement("throw new ArgumentNullException(\"env\");", 4),
                    new TranslatedStatement("", 0),
                    new TranslatedStatement(
                        string.Format("var {0} = env;", _envRefName.Name),
                        3
                    ),
                    new TranslatedStatement(
                        string.Format("var {0} = new {1}({2}, {3});", _outerRefName.Name, _outerClassName.Name, _supportRefName.Name, _envRefName.Name),
                        3
                    ),
                    new TranslatedStatement("", 0)
                });
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
                        3
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
                        3
                    ));
            }
            if (_outputType == OutputTypeOptions.Executable)
            {
                translatedStatements = translatedStatements
                    .AddRange(new[]
                    {
                        new TranslatedStatement("}", 2),
                        new TranslatedStatement("", 0)
                    })
                    .AddRange(new[]
                    {
                        new TranslatedStatement("public class " + _outerClassName.Name, 2),
                        new TranslatedStatement("{", 2),
                        new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " " + _supportRefName.Name + ";", 3),
                        new TranslatedStatement("private readonly " + _outerClassName.Name + " " + _outerRefName.Name + ";", 3),
                        new TranslatedStatement("private readonly " + _envClassName.Name + " " + _envRefName.Name + ";", 3),
                        new TranslatedStatement("public " + _outerClassName.Name + "(" + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " compatLayer, " + _envClassName.Name + " env)", 3),
                        new TranslatedStatement("{", 3),
                        new TranslatedStatement("if (compatLayer == null)", 4),
                        new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", 5),
                        new TranslatedStatement("if (env == null)", 4),
                        new TranslatedStatement("throw new ArgumentNullException(\"env\");", 5),
                        new TranslatedStatement(_supportRefName.Name + " = compatLayer;", 4),
                        new TranslatedStatement(_envRefName.Name + " = env;", 4),
                        new TranslatedStatement(_outerRefName.Name + " = this;", 4)
                    })
                    .AddRange(
                        explicitVariableDeclarationsFromWithOuterScope.Select(
                            v => new TranslatedStatement(
                                base.TranslateVariableInitialisation(v, ScopeLocationOptions.OutermostScope),
                                4
                            )
                        )
                    )
                    .Add(
                        new TranslatedStatement("}", 3)
                    );
                if (explicitVariableDeclarationsFromWithOuterScope.Any())
                {
                    translatedStatements = translatedStatements.Add(new TranslatedStatement("", 0));
                    translatedStatements = translatedStatements.AddRange(
                        explicitVariableDeclarationsFromWithOuterScope.Select(
                            v => new TranslatedStatement("public object " + _nameRewriter.GetMemberAccessTokenName(v.Name) + " { get; set; }", 3)
                        )
                    );
                }
            }
            foreach (var functionBlock in functionBlocks)
            {
                translatedStatements = translatedStatements.Add(new TranslatedStatement("", 1));
                translatedStatements = translatedStatements.AddRange(
                    Translate(
                        new NonNullImmutableList<ICodeBlock>(new[] { functionBlock }),
                        scopeAccessInformation,
                        3 // indentationDepth
                    ).TranslatedStatements
                );
            }
            if (_outputType == OutputTypeOptions.Executable)
            {
                translatedStatements = translatedStatements.AddRange(new[]
                {
                    new TranslatedStatement("}", 2),
                    new TranslatedStatement("", 0)
                });

                // This has to be generated after all of the Translate calls to ensure that the UndeclaredVariablesAccessed data for all
                // of the TranslationResults is available
                var allEnvironmentVariablesAccessed =
                    scopeAccessInformation.ExternalDependencies
                    .AddRange(
                        outerExecutableBlocksTranslationResult
                            .Add(classBlocksTranslationResult)
                            .UndeclaredVariablesAccessed
                    );
                translatedStatements = translatedStatements.AddRange(new[]
                {
                    new TranslatedStatement("public class " + _envClassName.Name, 2),
                    new TranslatedStatement("{", 2)
                });
                translatedStatements = translatedStatements.AddRange(
                    allEnvironmentVariablesAccessed
                        .Select(v => _nameRewriter.GetMemberAccessTokenName(v))
                        .Distinct()
                        .Select(v => new TranslatedStatement("public object " + v + " { get; set; }", 3)
                    )
                );
                translatedStatements = translatedStatements.Add(
                    new TranslatedStatement("}", 2)
                );
            }

            if (classBlocksTranslationResult.TranslatedStatements.Any())
            {
                translatedStatements = translatedStatements.Add(
                    new TranslatedStatement("", 0)
                );
                translatedStatements = translatedStatements.AddRange(
                    classBlocksTranslationResult.TranslatedStatements
                );
            }

            if (_outputType == OutputTypeOptions.Executable)
            {
                translatedStatements = translatedStatements.AddRange(new[]
                {
                    new TranslatedStatement("}", 1), // Close outer class
                    new TranslatedStatement("}", 0), // Close namespace
                });
            }
            return translatedStatements;
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
    }
}