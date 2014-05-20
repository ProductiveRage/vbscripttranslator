﻿using CSharpWriter.CodeTranslation.Extensions;
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
            if (scopeAccessInformation.ScopeLocation != ScopeLocationOptions.WithinFunctionOrProperty)
            {
                // This work is not required when not in a function - variable declarations for the outermost scope and within classes needs
                // different handling (in both cases, there will be properties on a class that will be set in the constructor - in the case
                // of the outermost scope, it is a class generated by this translation process, in the case of a VBScript class it will be
                // a translated class that represents that VBScript class)
                return translationResult;
            }
            return FlushUndeclaredVariableDeclarations(
                FlushExplicitVariableDeclarations(translationResult, indentationDepth),
                scopeAccessInformation.ScopeLocation,
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
			// This covers the DimStatement, PrivateVariableStatement and PublicVariableStatement (but not the ReDimStatement, which counts
            // as operating against an undeclared variable unless there is a corresponding Dim / Private / Public declaration)
			var explicitVariableDeclarationBlock = block as DimStatement;
			if (explicitVariableDeclarationBlock == null)
				return null;

            // These only need to result in additions to the ExplicitVariableDeclarations set of the translationResult, these will be
            // translated into the required form (varies depending upon whether the variable is defined within a class, function or
            // in outermost scope)
            return translationResult.Add(
                explicitVariableDeclarationBlock.Variables.Select(v => new VariableDeclaration(
                    v.Name,
                    (explicitVariableDeclarationBlock is PublicVariableStatement) ? VariableDeclarationScopeOptions.Public : VariableDeclarationScopeOptions.Private,
                    (v.Dimensions == null) ? null : v.Dimensions.Select(d => (uint)d.Value)
                ))
            );
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
            if (!(block is OptionExplicit))
                return null;

            // 2014-05-02 DWR: In order to support the absence of Option Explicit (the recommended way is to INCLUDE it but it is not compulsory),
            // there is support for tracking undeclared variables. If we're dealing with well-written VBScript with Option Explicit enabled then
            // there shouldn't be any places that undeclared variables are used and so ignoring Option Explicit SHOULDN'T cause any problems.
            // However, ignoring it does allow for a particularly awkward edge case around ReDim translation to be ignored. The following code:
            //
            //   Option Explicit
            //   If (False) Then
            //    ReDim a
            //   End If
            //   WScript.Echo TypeName(a)
            //
            // will result in a "Variable is undefined: 'a'" runtime error, though if "False" is replaced by "True" then it will write out
            // "Variant()" to the console. If "Option Explicit" is not included then the first case will result in "Empty" being written
            // out. This means that Option Explicit may result in variables being only "potentially declared", with its state not being
            // known until runtime. Handling this correctly would make the code translation significantly harder, since state about
            // which variables are and aren't currently declared would have to be maintained. This should be the only way in which
            // a script written with Option Explicit acts differently when translated (and since it is an edge case and indicates
            // poor practice even though Option Explicit has been used, I'm willing to live with the trade-off).
            _logger.Warning("Option Explicit is ignored by this translation process");
            return TranslationResult.Empty;
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

        protected TranslationResult TryToTranslateReDim(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
        {
            var reDimStatement = block as ReDimStatement;
            if (reDimStatement == null)
                return null;

            // Any variables that are referenced by the REDIM statement should be treated as if they HAVE been explicitly declared, so they will
            // be included in the returned TranslationResult's ExplicitVariableDeclarations set (any variables that have already been declared
            // need not be). Note that this means that code such as
            //
            //   REDIM a(0)
            //   DIM a
            //
            // will fail, but that is consistent with VBScript's behaviour (a compile time error). This is a situation where the DIM statement
            // does not effectively get hoisted to the top of the function. This behaviour happens regardless of the absence or presence of
            // OPTION EXPLICIT (since that can only cause run time errors and this is a "Name redefined" VBScript compilation error).
            var uninitialisedVariableDeclarationsToRecord = reDimStatement.Variables
                .Select(v => new
                {
                    SourceName = v.Name,
                    VariableDeclaration = new VariableDeclaration(
                        v.Name,
                        VariableDeclarationScopeOptions.Private,
                        null
                    )
                })
                .Where(
                    newVariable => !translationResult.ExplicitVariableDeclarations.Any(
                        existingVariable => existingVariable.Name.Content.Equals(newVariable.VariableDeclaration.Name.Content, StringComparison.OrdinalIgnoreCase)
                    )
                )
                .Where(newVariable =>
                    (scopeAccessInformation.ScopeDefiningParentIfAny == null) ||
                    !scopeAccessInformation.ScopeDefiningParentIfAny.Name.Content.Equals(newVariable.SourceName.Content, StringComparison.OrdinalIgnoreCase)
                );
            
            // These variables are now used to extend the current scopeAccessInformation, this is important for the translation below - any
            // variables that had not be declared should be treated as if they had been explicitly declared (this will affect the results of
            // calls to scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired; it will be able to correctly associated any previously-
            // undeclared variables with the local scope or outer most scope - depending upon whether we're within a function / property or
            // not - without updating the scopeAccessInformation, if we were in the outer most scope then the undeclared variables would be
            // identified as an "Environment References" entry, which we don't want).
            scopeAccessInformation = scopeAccessInformation.ExtendVariables(
                uninitialisedVariableDeclarationsToRecord
                    .Select(v => new ScopedNameToken(
                        v.SourceName.Content,
                        v.SourceName.LineIndex,
                        scopeAccessInformation.ScopeLocation
                    ))
                    .ToNonNullImmutableList()
            );

            var translatedReDimStatements = new NonNullImmutableList<TranslatedStatement>();
            var translatedContentFormat = reDimStatement.Preserve
                ? "{0} = {1}.RESIZEARRAY({0}, {2});"
                : "{0} = {1}.NEWARRAY({2});";
            foreach (var variable in reDimStatement.Variables)
            {
                var rewrittenVariableName = _nameRewriter(variable.Name).Name;
                string targetReference;
                if ((scopeAccessInformation.ScopeLocation == ScopeLocationOptions.WithinFunctionOrProperty)
                && variable.Name.Content.Equals(scopeAccessInformation.ScopeDefiningParentIfAny.Name.Content, StringComparison.OrdinalIgnoreCase))
                {
                    // REDIM statements can target the return value of a function - eg.
                    //
                    //   FUNCTION F1()
                    //    REDIM F1(0)
                    //    F1(0) = "Result"
                    //   END FUNCTION
                    //
                    // In this case, the "targetReference" value needs to be the name of the temporary value used as the function return
                    // value (the ScopeAccessInformation class should always have non-null references for the ScopeDefiningParentIfAny
                    // and ParentReturnValueNameIfAny properties if the ScopeLocation is WithinFunctionOrProperty).
                    targetReference = scopeAccessInformation.ParentReturnValueNameIfAny.Name;
                }
                else
                {
                    // If the target is not the function / property return value then we need to determine whether it's a local reference
                    // or within another scope. If it is currently undeclared, then this will return null if we're within a function or
                    // property or the name of the "Environment References" class if in the outer most scope. This is correct, we don't
                    // need to worry about the fact that any undeclared variables will be getting added to the ExplicitVariableDeclarations
                    // set of the return TranslationResult since they adding them to that set will result in them being declared as either
                    // local references or within the "Environment References", depending upon whether we're within a function / property
                    // or not - in other words, the exact same result.
                    var targetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenVariableName, _envRefName, _outerRefName, _nameRewriter);
                    targetReference = (targetContainer == null)
                        ? rewrittenVariableName
                        : (targetContainer.Name + "." + rewrittenVariableName);
                }

                var translatedArguments = new List<string>();
                foreach (var dimension in variable.Dimensions)
                {
                    var translatedArgumentDetails = _statementTranslator.Translate(
                        dimension,
                        scopeAccessInformation,
                        ExpressionReturnTypeOptions.Value
                    );
                    translatedArguments.Add(translatedArgumentDetails.TranslatedContent);
                    translationResult = translationResult.Add(translatedArgumentDetails.VariablesAccessed);
                }
                translatedReDimStatements = translatedReDimStatements.Add(
                    new TranslatedStatement(
                        string.Format(
                            translatedContentFormat,
                            targetReference,
                            _supportRefName.Name,
                            string.Join(", ", translatedArguments)
                        ),
                        indentationDepth
                    )
                );
            }

            return translationResult
                .Add(uninitialisedVariableDeclarationsToRecord.Select(v => v.VariableDeclaration))
                .Add(translatedReDimStatements);
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

        /// <summary>
        /// Within a function or property, this will be translated into a variable declaration AND initialisation (eg. "object v = null"
        /// or "object v = new object[1, 2]") but if the scope location is within a class (not within a function or property within a
        /// class, but directly within the class definition) or within the outermost scope then this will only initialise the value
        /// (eg. "v = null" or "v = new object[1, 2]") since the variable will be declared in a structure elsewhere (in an
        /// "OuterReferences" class, for example).
        /// </summary>
        protected string TranslateVariableInitialisation(VariableDeclaration variableDeclaration, ScopeLocationOptions scopeLocation)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");

            // For variables declared in the outermost scope or within functions, this could have been simplified such that a "Dim a()"
            // be rewritten as "object a = null;" and "a = _.NEWARRAY();" separately, which would mean that the VariableDeclaration
            // class need not have an IsArray method. But for statements within a class, such as "Private mValue()", this would not
            // be possible since the separate setter statement could not exist outside of a method.
            // 2014-04-24 DWR: Can't use NEWARRAY here since it relies upon an IProvideVBScriptCompatFunctionality instance, which
            // won't be available when initialising class members, so need to replicate the logic for an array initialisation that
            // doesn't have any dimensions here.
            // 2014-04-28 DWR: And can't set to null if there are array dimensions to set since class member array initialisations
            // will not be able to be set in a second pass so "Private a(1)" must be translated direct into "object a = new object[2]"
            // (noting that the dimension size is one more in C# since it is specifies the number of items in the array rather than
            // the upper bound).
            var rewrittenName = _nameRewriter.GetMemberAccessTokenName(variableDeclaration.Name);
            if (variableDeclaration.ConstantDimensionsIfAny == null)
            {
                return string.Format(
                    "{0}{1} = null;",
                    (scopeLocation == ScopeLocationOptions.WithinFunctionOrProperty) ? "object " : "",
                    rewrittenName
                );
            }
            else if (!variableDeclaration.ConstantDimensionsIfAny.Any())
            {
                return string.Format(
                    "{0}{1} = (object[])null;",
                    (scopeLocation == ScopeLocationOptions.WithinFunctionOrProperty) ? "object " : "",
                    rewrittenName
                );
            }
            return string.Format(
                "{0}{1} = new object[{2}];",
                rewrittenName,
                    (scopeLocation == ScopeLocationOptions.WithinFunctionOrProperty) ? "object " : "",
                string.Join(", ", variableDeclaration.ConstantDimensionsIfAny.Select(d => d + 1))
            );
        }

        /// <summary>
        /// This should only be called within a function or property since explicit variable declarations must be handled in special ways when within
        /// a class definition or in the outermost scope
        /// </summary>
        private TranslationResult FlushExplicitVariableDeclarations(TranslationResult translationResult, int indentationDepthForExplicitVariableDeclarations)
		{
			// TODO: Consider trying to insert the content after any comments or blank lines?
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
            if (indentationDepthForExplicitVariableDeclarations < 0)
                throw new ArgumentOutOfRangeException("indentationDepthForExplicitVariableDeclarations", "must be zero or greater");

            // While there may be duplicate references to undeclared variables, there may be only one explicit variable declaration for any variable
            // (VBScript will raise a "Name redefined" compile error if there are multiple DIM statements for the same variable - being a compile
            // error, it can not be avoided with On Error Resume Next so we should throw a translation exception)
            ThrowExceptionForDuplicateVariableDeclarationNames(translationResult.ExplicitVariableDeclarations);

            return new TranslationResult(
                translationResult.ExplicitVariableDeclarations
                    .Select(
                        v => new TranslatedStatement(TranslateVariableInitialisation(v, ScopeLocationOptions.WithinFunctionOrProperty), indentationDepthForExplicitVariableDeclarations)
                    )
                    .ToNonNullImmutableList()
				    .AddRange(translationResult.TranslatedStatements),
				new NonNullImmutableList<VariableDeclaration>(),
				translationResult.UndeclaredVariablesAccessed
			);
		}

        /// <summary>
        /// This will throw an exception for any duplicated variable name, which would have resulted in a VBScript "Name refined" compile error if
        /// present within the same scope. This uses a case-insensitive string comparison, to mimic VBScript (it does not consider variables that
        /// clash after being processed by the name rewriter since that would be a configuration issue, not an invalid-VBScript issue)
        /// </summary>
        protected void ThrowExceptionForDuplicateVariableDeclarationNames(NonNullImmutableList<VariableDeclaration> variableDeclarations)
        {
            if (variableDeclarations == null)
                throw new ArgumentNullException("variableDeclarations");

            var groupedVariableNames = variableDeclarations.GroupBy(v => v.Name.Content, StringComparer.OrdinalIgnoreCase);
            var firstDuplicateNameEntryIfAny = groupedVariableNames.FirstOrDefault(g => g.Count() > 1);
            if (firstDuplicateNameEntryIfAny == null)
                return;

            throw new ArgumentException(string.Format(
                "Multiple explicit variable declarations encountered within the same scope, this would be a VBScript compilation error - \"{0}\" on lines {1}",
                firstDuplicateNameEntryIfAny.Key,
                string.Join(", ", firstDuplicateNameEntryIfAny.Select(v => v.Name.LineIndex))
            ));
        }

        /// <summary>
        /// This should only be called within a function or property since undeclared variable declarations must be handled in special ways otherwise
        /// (added to an "ExternalReferences" class)
        /// </summary>
        private TranslationResult FlushUndeclaredVariableDeclarations(
            TranslationResult translationResult,
            ScopeLocationOptions scopeLocation,
            int indentationDepthForExplicitVariableDeclarations)
        {
            // TODO: Consider trying to insert the content after any comments or blank lines?
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");
            if (indentationDepthForExplicitVariableDeclarations < 0)
                throw new ArgumentOutOfRangeException("indentationDepthForExplicitVariableDeclarations", "must be zero or greater");

            // There could be multiple references to the same undeclared variable within any given scope and so there may be duplicate
            // entries in the UndeclaredVariablesAccessed set. Any duplicates are ignored (while the order, ignoring subsequent duplicate
            // values, is maintained).
            var rewrittenNamesAccountedFor = new HashSet<string>();
            var uniqueVariables = new List<VariableDeclaration>();
            foreach (var undeclaredVariable in translationResult.UndeclaredVariablesAccessed)
            {
                var rewrittenName = _nameRewriter(undeclaredVariable).Name;
                if (rewrittenNamesAccountedFor.Contains(rewrittenName))
                    continue;

                uniqueVariables.Add(
                    new VariableDeclaration(undeclaredVariable, VariableDeclarationScopeOptions.Private, null)
                );
                rewrittenNamesAccountedFor.Add(rewrittenName);
            }

            return new TranslationResult(
                uniqueVariables
                    .Select(v =>
                        new TranslatedStatement(
                            TranslateVariableInitialisation(v, ScopeLocationOptions.WithinFunctionOrProperty) + " /* Undeclared in source */",
                            indentationDepthForExplicitVariableDeclarations
                        )
                    )
                    .ToNonNullImmutableList()
                    .AddRange(translationResult.TranslatedStatements),
                translationResult.ExplicitVariableDeclarations,
                new NonNullImmutableList<NameToken>()
            );
        }
    }
}
