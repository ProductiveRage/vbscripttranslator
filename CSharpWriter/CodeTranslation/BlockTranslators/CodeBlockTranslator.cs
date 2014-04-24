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
    public abstract class CodeBlockTranslator
    {
        protected readonly CSharpName _supportRefName, _envClassName, _envRefName, _outerClassName, _outerRefName;
        protected readonly VBScriptNameRewriter _nameRewriter;
		protected readonly TempValueNameGenerator _tempNameGenerator;
        protected readonly ITranslateIndividualStatements _statementTranslator;
		private readonly ITranslateValueSettingsStatements _valueSettingStatementTranslator;
        private readonly ILogInformation _logger;
        protected CodeBlockTranslator(
            CSharpName supportRefName,
            CSharpName envClassName,
            CSharpName envRefName,
            CSharpName outerClassName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator,
			ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
        {
            if (supportRefName == null)
                throw new ArgumentNullException("supportRefName");
            if (envClassName == null)
                throw new ArgumentNullException("envClassName");
            if (envRefName == null)
                throw new ArgumentNullException("envRefName");
            if (outerClassName == null)
                throw new ArgumentNullException("outerClassName");
            if (outerRefName == null)
                throw new ArgumentNullException("outerRefName");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (tempNameGenerator == null)
                throw new ArgumentNullException("tempNameGenerator");
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
			if (valueSettingStatementTranslator == null)
				throw new ArgumentNullException("valueSettingStatementTranslator");
            if (logger == null)
                throw new ArgumentNullException("logger");

            _supportRefName = supportRefName;
            _envClassName = envClassName;
            _envRefName = envRefName;
            _outerClassName = outerClassName;
            _outerRefName = outerRefName;
            _nameRewriter = nameRewriter;
            _tempNameGenerator = tempNameGenerator;
            _statementTranslator = statementTranslator;
			_valueSettingStatementTranslator = valueSettingStatementTranslator;
            _logger = logger;
        }

		/// <summary>
		/// This should return null if it is unable to process the specified block. It should raise an exception for any null arguments. The returned value
        /// (where non-null) should overwrite the input translationResult in the caller's scope (it should not be added to it).
		/// </summary>
		protected delegate TranslationResult BlockTranslationAttempter(
			TranslationResult translationResult,
			ICodeBlock block,
			ScopeAccessInformation scopeAccessInformation,
			int indentationDepth
		);

		protected TranslationResult TranslateCommon(
			NonNullImmutableList<BlockTranslationAttempter> translators,
			NonNullImmutableList<ICodeBlock> blocks,
			ScopeAccessInformation scopeAccessInformation,
			int indentationDepth)
        {
			if (translators == null)
				throw new ArgumentNullException("translators");
            if (blocks == null)
                throw new ArgumentNullException("block");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			// TODO: Going to need to incorporate On Error Resume Next / Goto 0 handling outside of the rest of the process, requiring additional data
			// in the scopeAccessInformation type?

            var translationResult = TranslationResult.Empty;
			foreach (var block in blocks)
            {
				var hasBlockBeenTranslated = false;
				foreach (var translator in translators)
				{
					var blockTranslationResult = translator(translationResult, block, scopeAccessInformation, indentationDepth);
					if (blockTranslationResult == null)
						continue;

                    // Note: translationResult is set to blockTranslationResult rather than translationResult.Add(blockTranslationResult) since the
                    // translationResult is passed into the translator delegate above and that is reponsible for handling any combination work
                    // required
                    translationResult = blockTranslationResult;
					hasBlockBeenTranslated = true;
                    break;
				}
				if (!hasBlockBeenTranslated)
					throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
            }

            // If the current parent construct doesn't affect scope (like IF and WHILE and unlike CLASS and FUNCTION) then the translationResult
            // can be returned directly and the nearest construct that does affect scope will be responsible for translating any explicit
            // variable declarations into translated statements. If the "scope-defining parent" is the outermost scope then parentIfAny
            // will be null but that's ok since explicit variable declarations aren't flushed in the same way in that scope, they are
            // added to an "outer" or "GlobalReferences" class.
            var scopeDefiningParent = scopeAccessInformation.ParentIfAny as IDefineScope;
			if (scopeDefiningParent == null)
                return translationResult;

            // Explicitly-declared variable declarations need to be translated into C# definitions here (hoisted to the top of the function), as
            // do any undeclared variables (in VBScript if an undeclared variable is used within a function or property body then that variable
            // is treated as being local to the function or property)
            return FlushUndeclaredVariableDeclarations(            
                FlushExplicitVariableDeclarations(
                    translationResult,
                    indentationDepth
                ),
                indentationDepth
            );
        }

		protected TranslationResult TryToTranslateBlankLine(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
            return (block is BlankLine) ? translationResult.Add(new TranslatedStatement("", indentationDepth)) : null;
		}

		protected TranslationResult TryToTranslateClass(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var classBlock = block as ClassBlock;
			if (classBlock == null)
				return null;

			var codeBlockTranslator = new ClassBlockTranslator(
                _supportRefName,
                _envClassName, 
                _envRefName,
                _outerClassName,
                _outerRefName,
                _nameRewriter,
                _tempNameGenerator,
                _statementTranslator,
                _valueSettingStatementTranslator,
                _logger
            );
			return translationResult.Add(
				codeBlockTranslator.Translate(
					classBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateComment(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var commentBlock = block as CommentStatement;
			if (commentBlock == null)
				return null;

			var translatedCommentContent = "//" + commentBlock.Content;
			if (block is InlineCommentStatement)
			{
				var lastTranslatedStatement = translationResult.TranslatedStatements.LastOrDefault();
				if ((lastTranslatedStatement != null) && (lastTranslatedStatement.Content != ""))
				{
					translationResult = new TranslationResult(
						translationResult.TranslatedStatements
							.RemoveLast()
							.Add(new TranslatedStatement(
								lastTranslatedStatement.Content + " " + translatedCommentContent,
								lastTranslatedStatement.IndentationDepth
							)),
						translationResult.ExplicitVariableDeclarations,
						translationResult.UndeclaredVariablesAccessed
					);
					return translationResult;
				}
			}

			return translationResult.Add(
				new TranslatedStatement(translatedCommentContent, indentationDepth)
			);
		}

		protected TranslationResult TryToTranslateDim(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
            // TODO: Within a function, REDIM may be used on the function name to initialise an array for the return value. This is not dealt with yet.

			// This covers the DimStatement, ReDimStatement, PrivateVariableStatement and PublicVariableStatement
			var explicitVariableDeclarationBlock = block as DimStatement;
			if (explicitVariableDeclarationBlock == null)
				return null;

			// Ensure that all DIM'd variables are recorded as explicit variable declarations. It is important that it be recorded if these variables
            // are declared as arrays - eg. "Dim a()" or "Dim a(0)" or "Private a()" - since this affects whether they are initialised as "null" or
            // "(object[])null" later on (this distinction is important since the IsArray method returns true for a if "Dim a()" is used but false
            // if "Dim a" is). This array/not-array has to be set in a single statement so that class member variables can be set correctly (in the
            // outermost scope of within a function, all DIM statements could be translated into a null-value initialisation followed by a separate
            // setting to a null array if it has dimensions, but this is not possible in a class member initialisation, as it must happen within
            // one C# statement). One more oddity is that VBScript will accept a REDIM for a variable that has not already had a DIM statement
            // for it, even if Option Explicit is specified. To deal with this, the REDIM target variables will be declared as explicit variable
            // declarations here and then processed in a separate pass below.
            // Note: There is no need to worry about specifying the same variable multiple times in the ExplicitVariableDeclarations data, that
            // will be dealt with when the set is translated into C# code later on.
            translationResult = translationResult.Add(
				explicitVariableDeclarationBlock.Variables.Select(v =>
					new VariableDeclaration(
						v.Name,
						(explicitVariableDeclarationBlock is PublicVariableStatement)
							? VariableDeclarationScopeOptions.Public
							: VariableDeclarationScopeOptions.Private,
						(v.Dimensions != null) // If the Dimensions set is non-null then this is an array type
					)
				)
			);

            // Any (RE)DIM statements with dimensions are now translated into array initialisers in the translated output (the above work to extend
            // the ExplicitVariableDeclarations set will effectively hoist the variable declarations to the top of the scope whereas these statements
            // will remain "in order"). In VBScript, a DIM statement may only have constant dimensions, unlike REDIM. Not meeting this would result in
            // a compile time error and so this translation code need not worry about it (the assumption is that we are dealing with valid VBScript).
            var isAnArrayExtension = (explicitVariableDeclarationBlock is ReDimStatement) && ((ReDimStatement)explicitVariableDeclarationBlock).Preserve;
            var callFormat = isAnArrayExtension
                ? "{0} = {1}.EXTENDARRAY({0}, {2})"
                : "{0} = {1}.NEWARRAY({2})";
            foreach (var arrayDeclaration in explicitVariableDeclarationBlock.Variables.Where(v => (v.Dimensions ?? new Expression[0]).Any()))
            {
                var rewrittenVariableName = _nameRewriter(arrayDeclaration.Name).Name;
                var nameOfTargetContainerIfRequired = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(
                    rewrittenVariableName,
                    _envRefName,
                    _outerRefName,
                    _nameRewriter
                );
                if (nameOfTargetContainerIfRequired != null)
                    rewrittenVariableName = nameOfTargetContainerIfRequired.Name + "." + rewrittenVariableName;

                var translatedArrayDimensionExpressionDetails = arrayDeclaration.Dimensions.Select(d =>
                    _statementTranslator.Translate(d, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified)
                );
                translationResult = translationResult
                    .Add(translatedArrayDimensionExpressionDetails.SelectMany(t => t.VariablesAccessed))
                    .Add(new TranslatedStatement(
                        string.Format(
                            callFormat,
                            rewrittenVariableName,
                            _supportRefName.Name,
                            string.Join(
                                ", ",
                                translatedArrayDimensionExpressionDetails.Select(t => t.TranslatedContent)
                            )
                        ),
                        indentationDepth
                    ));
            }
            return translationResult;
		}

		protected TranslationResult TryToTranslateDo(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var doBlock = block as DoBlock;
			if (doBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateExit(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var exitStatement = block as ExitStatement;
			if (exitStatement == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateFor(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var forBlock = block as ForBlock;
			if (forBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateForEach(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var forEachBlock = block as ForEachBlock;
			if (forEachBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateFunction(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var functionBlock = block as AbstractFunctionBlock;
			if (functionBlock == null)
				return null;

			var codeBlockTranslator = new FunctionBlockTranslator(
                _supportRefName,
                _envClassName, 
                _envRefName,
                _outerClassName,
                _outerRefName,
                _nameRewriter,
                _tempNameGenerator,
                _statementTranslator,
                _valueSettingStatementTranslator,
                _logger
            );
			return translationResult.Add(
				codeBlockTranslator.Translate(
					functionBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateIf(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var ifBlock = block as IfBlock;
			if (ifBlock == null)
				return null;

			var codeBlockTranslator = new IfBlockTranslator(
                _supportRefName,
                _envClassName, 
                _envRefName,
                _outerClassName,
                _outerRefName,
                _nameRewriter,
                _tempNameGenerator,
                _statementTranslator,
                _valueSettingStatementTranslator,
                _logger
            );
			return translationResult.Add(
				codeBlockTranslator.Translate(
					ifBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateOptionExplicit(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			return (block is OptionExplicit) ? TranslationResult.Empty : null;
		}

		protected TranslationResult TryToTranslateProperty(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var propertyBlock = block as PropertyBlock;
			if (propertyBlock == null)
				return null;

			// Note: Check for Default on the Get (the only place it's valid) before rendering Let or Set
			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateRandomize(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var randomizeStatement = block as RandomizeStatement;
			if (randomizeStatement == null)
				return null;

			// Note: See http://msdn.microsoft.com/en-us/library/e566zd96(v=vs.84).aspx when implementing RND
			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateSelect(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var selectBlock = block as SelectBlock;
			if (selectBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateStatementOrExpression(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			// This covers Statement and Expression instances as Expression inherits from Statement
			var statementBlock = block as Statement;
			if (statementBlock == null)
				return null;

            // TODO: Differentiate between undeclared variables and functions (and classes?)
            var translatedStatementContentDetails = _statementTranslator.Translate(statementBlock, scopeAccessInformation);
            var undeclaredVariables = translatedStatementContentDetails.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(_nameRewriter(v).Name, _nameRewriter));
            foreach (var undeclaredVariable in undeclaredVariables)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
			return
                translationResult.Add(
				    new TranslatedStatement(
                        translatedStatementContentDetails.TranslatedContent + ";",
					    indentationDepth
				    )
			    )
				.Add(undeclaredVariables);
		}

		protected TranslationResult TryToTranslateValueSettingStatement(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var valueSettingStatement = block as ValueSettingStatement;
			if (valueSettingStatement == null)
				return null;

            // TODO: Differentiate between undeclared variables and functions (and classes?)
            var translatedValueSettingStatementContentDetails = _valueSettingStatementTranslator.Translate(valueSettingStatement, scopeAccessInformation);
            var undeclaredVariables = translatedValueSettingStatementContentDetails.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(_nameRewriter(v).Name, _nameRewriter));
            foreach (var undeclaredVariable in undeclaredVariables)
				_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
			return
				translationResult.Add(
					new TranslatedStatement(
						translatedValueSettingStatementContentDetails.TranslatedContent + ";",
						indentationDepth
					)
				)
				.Add(undeclaredVariables);
		}

        protected string TranslateVariableDeclaration(VariableDeclaration variableDeclaration)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");

            // For variables declared in the outermost scope or within functions, this could have been simplified such that a "Dim a()"
            // be rewritten as "object a = null;" and "a = _.NEWARRAY();" separately, which would mean that the VariableDeclaration
            // class need not have an IsArray method. But for statements within a class, such as "Private mValue()", this would not
            // be possible since the separate setter statement could not exist outside of a method.
            var rewrittenName = _nameRewriter.GetMemberAccessTokenName(variableDeclaration.Name);
            if (variableDeclaration.IsArray)
            {
                return string.Format(
                    "object {0} = {1}.NEWARRAY();",
                    rewrittenName,
                    _supportRefName.Name
                );

            }
            return string.Format(
                "object {0} = null;",
                rewrittenName
            );
        }

        private TranslationResult FlushExplicitVariableDeclarations(TranslationResult translationResult, int indentationDepthForExplicitVariableDeclarations)
		{
			// TODO: Consider trying to insert the content after any comments or blank lines?
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
            if (indentationDepthForExplicitVariableDeclarations < 0)
                throw new ArgumentOutOfRangeException("indentationDepthForExplicitVariableDeclarations", "must be zero or greater");

            var uniqueVariableDeclarations = RemoveDuplicateVariableNames(translationResult.ExplicitVariableDeclarations);
            return new TranslationResult(
                uniqueVariableDeclarations
					.Select(v =>
                        new TranslatedStatement(TranslateVariableDeclaration(v), indentationDepthForExplicitVariableDeclarations)
					)
                    .ToNonNullImmutableList()
					.AddRange(translationResult.TranslatedStatements),
				new NonNullImmutableList<VariableDeclaration>(),
				translationResult.UndeclaredVariablesAccessed
			);
		}

        private TranslationResult FlushUndeclaredVariableDeclarations(TranslationResult translationResult, int indentationDepthForExplicitVariableDeclarations)
        {
            // TODO: Consider trying to insert the content after any comments or blank lines?
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepthForExplicitVariableDeclarations < 0)
                throw new ArgumentOutOfRangeException("indentationDepthForExplicitVariableDeclarations", "must be zero or greater");

            var uniqueVariableDeclarations = RemoveDuplicateVariableNames(
                translationResult.UndeclaredVariablesAccessed.Select(v =>
                    new VariableDeclaration(v, VariableDeclarationScopeOptions.Private, false)
                )
                .ToNonNullImmutableList()
            );
            return new TranslationResult(
                uniqueVariableDeclarations
                    .Select(v =>
                        new TranslatedStatement(
                            TranslateVariableDeclaration(v) + " /* Undeclared in source */",
                            indentationDepthForExplicitVariableDeclarations
                        )
                    )
                    .ToNonNullImmutableList()
                    .AddRange(translationResult.TranslatedStatements),
                translationResult.ExplicitVariableDeclarations,
                new NonNullImmutableList<NameToken>()
            );
        }

        private NonNullImmutableList<VariableDeclaration> RemoveDuplicateVariableNames(NonNullImmutableList<VariableDeclaration> variableDeclarations)
        {
            if (variableDeclarations == null)
                throw new ArgumentNullException("variableDeclarations");

            var rewrittenNameToIsArrayLookup = new Dictionary<string, bool>();
            foreach (var variableDeclaration in variableDeclarations)
            {
                var rewrittenName = _nameRewriter.GetMemberAccessTokenName(variableDeclaration.Name);
                if (!rewrittenNameToIsArrayLookup.ContainsKey(rewrittenName))
                {
                    rewrittenNameToIsArrayLookup.Add(rewrittenName, variableDeclaration.IsArray);
                    continue;
                }
                rewrittenNameToIsArrayLookup[rewrittenName] = rewrittenNameToIsArrayLookup[rewrittenName] || variableDeclaration.IsArray;
            }

            var nonDuped = new List<VariableDeclaration>();
            var rewrittenNamesAccountedFor = new HashSet<string>();
            foreach (var variableDeclaration in variableDeclarations)
            {
                var rewrittenName = _nameRewriter.GetMemberAccessTokenName(variableDeclaration.Name);
                if (rewrittenNamesAccountedFor.Contains(rewrittenName))
                    continue;

                // If this variable is declared later on in the data as an array then ensure that it is recorded as an array reference
                // now since the later declaration will be ignored as a duplicate
                nonDuped.Add(new VariableDeclaration(
                    variableDeclaration.Name,
                    variableDeclaration.Scope,
                    rewrittenNameToIsArrayLookup[rewrittenName]
                ));
                rewrittenNamesAccountedFor.Add(rewrittenName);
            }
            return nonDuped.ToNonNullImmutableList();
        }
    }
}
