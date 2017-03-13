using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions;
using VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation;
using VBScriptTranslator.CSharpWriter.Lists;
using VBScriptTranslator.CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.BlockTranslators
{
	public abstract class CodeBlockTranslator
	{
		protected readonly CSharpName _supportRefName, _envClassName, _envRefName, _outerClassName, _outerRefName;
		protected readonly VBScriptNameRewriter _nameRewriter;
		protected readonly TempValueNameGenerator _tempNameGenerator;
		private readonly ITranslateIndividualStatements _statementTranslator;
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
			// variable declarations into translated statements.
			var scopeDefiningParent = scopeAccessInformation.Parent as IDefineScope;
			if (scopeDefiningParent == null)
				return translationResult;

			// Explicitly-declared variable declarations need to be translated into C# definitions here (hoisted to the top of the function), as
			// do any undeclared variables (in VBScript if an undeclared variable is used within a function or property body then that variable
			// is treated as being local to the function or property)
			if (scopeAccessInformation.ScopeDefiningParent.Scope != ScopeLocationOptions.WithinFunctionOrPropertyOrWith)
			{
				// This work is not required when not in a function - variable declarations for the outermost scope and within classes needs
				// different handling (in both cases, there will be properties on a class that will be set in the constructor - in the case
				// of the outermost scope, it is a class generated by this translation process, in the case of a VBScript class it will be
				// a translated class that represents that VBScript class)
				return translationResult;
			 }
			return FlushExplicitVariableDeclarations(
				FlushUndeclaredVariableDeclarations(translationResult, scopeAccessInformation.ScopeDefiningParent.Scope, indentationDepth),
				indentationDepth
			);
		}

		protected TranslationResult TryToTranslateBlankLine(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			// We don't have any information about where blank lines comes from as the code blocks don't contain this information directly
			// and there are no tokens in the BlankLine class to infer the information from. So we'll have to leave it as zero (it is
			// documented on TranslatedStatement that some line index values will be approximate or inaccurate in some cases).
			return (block is BlankLine) ? translationResult.Add(new TranslatedStatement("", indentationDepth, lineIndexOfStatementStartInSource: 0)) : null;
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
								lastTranslatedStatement.IndentationDepth,
								commentBlock.LineIndex
							)),
						translationResult.ExplicitVariableDeclarations,
						translationResult.UndeclaredVariablesAccessed
					);
					return translationResult;
				}
			}

			return translationResult.Add(
				new TranslatedStatement(translatedCommentContent, indentationDepth, commentBlock.LineIndex)
			);
		}

		private TranslationResult TryToTranslateConst(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var constStatement = block as ConstStatement;
			if (constStatement == null)
				return null;

			// Check for another CONST that appeared earlier in this block that specified the same variable(s) - this requires a "Name refined" compile / translation error.
			// We don't check for DIM statements here since the TryToTranslateDim functions deals with clashes of variable names in DIMs and CONSTS, so here we can pretend
			// that DIMs don't exist.
			// ^ TODO: No, no, no (doesn't catch the ReDim-then-Const case)
			var dimVariableNamesInCurrentScope = scopeAccessInformation.ScopeDefiningParent.GetAllNestedBlocks()
				.TakeWhile(b => b != block)
				.OfType<BaseDimStatement>()
				.SelectMany(dim => dim.Variables)
				.Select(v => v.Name);
			var constNamesInCurrentScope = scopeAccessInformation.ScopeDefiningParent.GetAllNestedBlocks()
				.TakeWhile(b => b != block)
				.OfType<ConstStatement>()
				.SelectMany(c => c.Values)
				.Select(v => v.Name);
			var previouslyDeclaredVariableNamesInCurrentScope = dimVariableNamesInCurrentScope.Concat(constNamesInCurrentScope);
			var firstVariableAlreadyDeclaredInTheCurrentScopeIfAny = previouslyDeclaredVariableNamesInCurrentScope.FirstOrDefault(
				previouslyDeclaredVariable => constStatement.Values.Any(v => _nameRewriter.AreNamesEquivalent(v.Name, previouslyDeclaredVariable))
			);
			if (firstVariableAlreadyDeclaredInTheCurrentScopeIfAny != null)
				throw new NameRedefinedException(firstVariableAlreadyDeclaredInTheCurrentScopeIfAny);

			// The CONST value-setting statements need to inserted at the top of the TranslationResult since a CONST is effectively hoisted (WITH its value) to the top of
			// its current scope. Note that this won't result in it being inserted before the explicit variable declarations, which is good. To illustrate:
			//
			//   FUNCTION F1()
			//     F2(cn1)
			//     CONST cn1 = 123
			//   END FUNCTION
			//
			// is translated into something like
			//
			//   public object f1()
			//   {
			//     object retVal1 = null;
			//     object cn1 = null;
			//     cn1 = 123;
			//     _.CALL(this, "F2", _.ARGS.Val(cn1));
			//     return retVal1;
			//   }
			//
			// The value "cn1" is set to 123 immediately, before any other processing (ie. the call to F2) occurs.
			return new TranslationResult(
				translationResult.TranslatedStatements.Insert(
					constStatement.Values.Select(value =>
						new TranslatedStatement(
							string.Format(
								"{0}{1} = {2};",
								(scopeAccessInformation.ScopeDefiningParent.Scope == ScopeLocationOptions.OutermostScope) ? (_outerRefName.Name + ".") : "",
								_nameRewriter.GetMemberAccessTokenName(value.Name),
								value.Value.Content
							),
							indentationDepth,
							value.Name.LineIndex
						)
					),
					0
				),
				translationResult.ExplicitVariableDeclarations.AddRange(
					constStatement.Values.Select(v => new VariableDeclaration(
						v.Name,
						VariableDeclarationScopeOptions.Public, // There are no private CONST statements so this is public by default
						constantDimensionsIfAny: null // This does not apply to CONST statements, they may never be arrays
					))
				),
				translationResult.UndeclaredVariablesAccessed
			);
		}

		protected TranslationResult TryToTranslateDim(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			// This covers the DimStatement, PrivateVariableStatement and PublicVariableStatement (but not the ReDimStatement, which counts
			// as operating against an undeclared variable unless there is a corresponding Dim / Private / Public declaration)
			var explicitVariableDeclarationBlock = block as DimStatement;
			if (explicitVariableDeclarationBlock == null)
				return null;

			// If a DIM statement exists after another DIM or REDIM for the same variable within the same scope, then a "Name redefined" compile
			// error would be raised by the VBScript compiler. Note that these do not have to occur on the same executation path - eg.
			//
			//    If (True) Then
			//        ReDim a(0)
			//    Else
			//        Dim a
			//    End If
			//
			// will result in an error since there is a REDIM for "a" before the DIM for the same variable, even though no one code path will
			// pass through both statements. This is because a REDIM will be considered by the interpreter to register an implicit variable
			// declaration within the current scope for the target and so the DIM is then considered to be a redefinition. The follow variation
			// will not result in a "Name redefined" error -
			//
			//    If (True) Then
			//        Dim a
			//    Else
			//        ReDim a(0)
			//    End If
			//
			// This is because it is a normal use case for a REDIM to target an already-declared variable reference.
			var dimVariableNamesInCurrentScope = scopeAccessInformation.ScopeDefiningParent.GetAllNestedBlocks()
				.TakeWhile(b => b != block)
				.OfType<BaseDimStatement>()
				.SelectMany(dim => dim.Variables)
				.Select(v => v.Name);
			var constNamesInCurrentScope = scopeAccessInformation.ScopeDefiningParent.GetAllNestedBlocks()
				.TakeWhile(b => b != block)
				.OfType<ConstStatement>()
				.SelectMany(c => c.Values)
				.Select(v => v.Name);
			var previouslyDeclaredVariableNamesInCurrentScope = dimVariableNamesInCurrentScope.Concat(constNamesInCurrentScope);
			var firstVariableAlreadyDeclaredInTheCurrentScopeIfAny = previouslyDeclaredVariableNamesInCurrentScope.FirstOrDefault(
				previouslyDeclaredVariable => explicitVariableDeclarationBlock.Variables.Any(v => _nameRewriter.AreNamesEquivalent(v.Name, previouslyDeclaredVariable))
			);
			if (firstVariableAlreadyDeclaredInTheCurrentScopeIfAny != null)
				throw new NameRedefinedException(firstVariableAlreadyDeclaredInTheCurrentScopeIfAny);

			return translationResult.AddExplicitVariableDeclarations(
				explicitVariableDeclarationBlock.Variables.Select(v => new VariableDeclaration(
					v.Name,
					// Dim and Public keywords = Public, Private keyword = Private
					(explicitVariableDeclarationBlock is PrivateVariableStatement) ? VariableDeclarationScopeOptions.Private : VariableDeclarationScopeOptions.Public,
					(v.Dimensions == null) ? null : v.Dimensions.Select(d => (uint)d.Value)
				))
			);
		}

		private TranslationResult TryToTranslateDo(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var doBlock = block as DoBlock;
			if (doBlock == null)
				return null;

			var codeBlockTranslator = new DoBlockTranslator(
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
					doBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		private TranslationResult TryToTranslateErase(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var eraseStatement = block as EraseStatement;
			if (eraseStatement == null)
				return null;

			var codeBlockTranslator = new EraseStatementTranslator(
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
				codeBlockTranslator.Translate(eraseStatement, scopeAccessInformation, indentationDepth)
			);
		}

		private TranslationResult TryToTranslateExit(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var exitStatement = block as ExitStatement;
			if (exitStatement == null)
				return null;

			// EXIT FUNCTION, PROPERTY or SUB statements are simple enough - if the current scope has a return value being maintained (FUNCTION
			// and PROPERTY will, SUB won't) then return that - otherwise just return (with a value). Some validation is done that the EXIT is
			// acceptable - eg. there isn't an EXIT SUB within a FUNCTION. Then, if there is error handling enabled within the the current
			// scope, the error handler token needs to be released as we are about to leave this scope (this is also done at the end of
			// the FunctionBlockTranslator's work, but since we're exiting early and not getting to that, we need to do it here too).
			bool isValidatedFunctionTypeExit;
			if (exitStatement.StatementType == ExitStatement.ExitableStatementType.Function)
			{
				if ((scopeAccessInformation.ScopeDefiningParent as FunctionBlock) == null)
					throw new ArgumentException("Encountered EXIT FUNCTION that was not within a function");
				isValidatedFunctionTypeExit = true;
			}
			else if (exitStatement.StatementType == ExitStatement.ExitableStatementType.Property)
			{
				if ((scopeAccessInformation.ScopeDefiningParent as PropertyBlock) == null)
					throw new ArgumentException("Encountered EXIT PROPERTY that was not within a property");
				isValidatedFunctionTypeExit = true;
			}
			else if (exitStatement.StatementType == ExitStatement.ExitableStatementType.Sub)
			{
				if ((scopeAccessInformation.ScopeDefiningParent as SubBlock) == null)
					throw new ArgumentException("Encountered EXIT SUB that was not within a sub");
				isValidatedFunctionTypeExit = true;
			}
			else
				isValidatedFunctionTypeExit = false;
			if (isValidatedFunctionTypeExit)
			{
				if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
				{
					translationResult = translationResult.Add(new TranslatedStatement(
						string.Format(
							"{0}.RELEASEERRORTRAPPINGTOKEN({1});",
							_supportRefName.Name,
							scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
						),
						indentationDepth,
						exitStatement.LineIndex
					));
				}
				return translationResult.Add(
					new TranslatedStatement(
						string.Format(
							"return{0};",
							(scopeAccessInformation.ParentReturnValueNameIfAny == null) ? "" : (" " + scopeAccessInformation.ParentReturnValueNameIfAny.Name)
						),
						indentationDepth,
						exitStatement.LineIndex
					)
				);
			}

			// For EXIT DO and EXIT FOR, we need to break out of the current structure, but that might not be enough. If we're within a FOR loop within a
			// DO loop and an EXIT DO is encountered then we need to break out of the FOR loop and then break out of the DO loop as well. The scope
			// information's StructureExitPoints allows us to do that, we can set the appropriate "exit-early" flag and then break out (the FOR
			// and DO translation implementations are responsible for checking the exit-early flags).
			CSharpName exitEarlyFlagForValidatedLoopTypeExitIfAny;
			if (exitStatement.StatementType == ExitStatement.ExitableStatementType.Do)
			{
				var correspondingExitableStructureDetails = scopeAccessInformation.StructureExitPoints.LastOrDefault(
					e => e.StructureType == ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions.Do
				);
				if (correspondingExitableStructureDetails == null)
					throw new ArgumentException("Encountered EXIT DO that was not within a do loop");
				exitEarlyFlagForValidatedLoopTypeExitIfAny = correspondingExitableStructureDetails.ExitEarlyBooleanNameIfAny;
			}
			else if (exitStatement.StatementType == ExitStatement.ExitableStatementType.For)
			{
				var correspondingExitableStructureDetails = scopeAccessInformation.StructureExitPoints.LastOrDefault(
					e => e.StructureType == ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions.For
				);
				if (correspondingExitableStructureDetails == null)
					throw new ArgumentException("Encountered EXIT FOR that was not within a for loop");
				exitEarlyFlagForValidatedLoopTypeExitIfAny = correspondingExitableStructureDetails.ExitEarlyBooleanNameIfAny;
			}
			else
				throw new ArgumentException("Unsupported ExitableStatementType: " + exitStatement.StatementType);
			if (exitEarlyFlagForValidatedLoopTypeExitIfAny != null)
			{
				translationResult = translationResult.Add(new TranslatedStatement(
					exitEarlyFlagForValidatedLoopTypeExitIfAny.Name + " = true;",
					indentationDepth,
					exitStatement.LineIndex
				));
			}
			return translationResult.Add(new TranslatedStatement(
				"break;",
				indentationDepth,
				exitStatement.LineIndex
			));
		}

		private TranslationResult TryToTranslateFor(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var forBlock = block as ForBlock;
			if (forBlock == null)
				return null;

			var codeBlockTranslator = new ForBlockTranslator(
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
					forBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		private TranslationResult TryToTranslateForEach(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var forEachBlock = block as ForEachBlock;
			if (forEachBlock == null)
				return null;

			var codeBlockTranslator = new ForEachBlockTranslator(
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
					forEachBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateFunctionPropertyOrSub(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
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

		private TranslationResult TryToTranslateIf(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
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

		private TranslationResult TryToTranslateOnErrorResumeNext(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var onErrorResumeNextBlock = block as OnErrorResumeNext;
			if (onErrorResumeNextBlock == null)
				return null;

			if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
				throw new ArgumentException("The ScopeAccessInformation's ErrorRegistrationTokenIfAny may not be null when the scope contains OnErrorResumeNext");

			// Note: Any time an "On Error Resume Next" statement is encountered, any current error information is cleared
			return translationResult.Add(new TranslatedStatement(
				string.Format(
					"{0}.STARTERRORTRAPPINGANDCLEARANYERROR({1});",
					_supportRefName.Name,
					scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
				),
				indentationDepth,
				onErrorResumeNextBlock.LineIndex
			));
		}

		private TranslationResult TryToTranslateOnErrorGotoZero(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var onErrorGotoZeroBlock = block as OnErrorGoto0;
			if (onErrorGotoZeroBlock == null)
				return null;

			// If this is a "On Error Goto 0" statement within a scope that does not contain an "On Error Resume Next" then it's redundant statement
			// and doesn't need to interact with the enable-disable-error-trapping-for-this-token logic, it only needs to clear any error
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
			{
				_logger.Warning("Ignoring ON ERROR GOTO 0 within a scope that contains no ON ERROR RESUME NEXT (line " + (onErrorGotoZeroBlock.LineIndex + 1) + ")");
				return translationResult.Add(new TranslatedStatement(
					_supportRefName.Name + ".CLEARANYERROR();",
					indentationDepth,
					onErrorGotoZeroBlock.LineIndex
				));
			}

			return translationResult.Add(new TranslatedStatement(
				string.Format(
					"{0}.STOPERRORTRAPPINGANDCLEARANYERROR({1});",
					_supportRefName.Name,
					scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
				),
				indentationDepth,
				onErrorGotoZeroBlock.LineIndex
			));
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

		private TranslationResult TryToTranslateRandomize(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var randomizeStatement = block as RandomizeStatement;
			if (randomizeStatement == null)
				return null;

			var translatedRandomizeStatements = new NonNullImmutableList<TranslatedStatement>();

			var mayRequireErrorWrapping = scopeAccessInformation.MayRequireErrorWrapping(block) && (randomizeStatement.SeedIfAny != null);
			if (mayRequireErrorWrapping)
			{
				translatedRandomizeStatements = translatedRandomizeStatements.Add(
					new TranslatedStatement(
						GetHandleErrorContent(scopeAccessInformation.ErrorRegistrationTokenIfAny),
						indentationDepth,
						randomizeStatement.SeedIfAny.Tokens.First().LineIndex
					)
				);
			}

			string translatedSeedIfAny;
			if (randomizeStatement.SeedIfAny == null)
				translatedSeedIfAny = null;
			else
			{
				var translatedSeedExpression = _statementTranslator.Translate(randomizeStatement.SeedIfAny, scopeAccessInformation, ExpressionReturnTypeOptions.Value, _logger.Warning);
				var undeclaredVariables = translatedSeedExpression.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
				foreach (var undeclaredVariable in undeclaredVariables)
					_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
				translationResult = translationResult.AddUndeclaredVariables(undeclaredVariables);
				translatedSeedIfAny = translatedSeedExpression.TranslatedContent;
			}

			translatedRandomizeStatements = translatedRandomizeStatements.Add(new TranslatedStatement(
				string.Format(
					"{0}.RANDOMIZE({1});",
					_supportRefName.Name,
					translatedSeedIfAny
				),
				indentationDepth,
				randomizeStatement.LineIndex
			));

			if (mayRequireErrorWrapping)
			{
				translatedRandomizeStatements = translatedRandomizeStatements.Add(
					new TranslatedStatement(
						"});",
						indentationDepth,
						randomizeStatement.SeedIfAny.Tokens.First().LineIndex
					)
				);
			}

			return translationResult.Add(translatedRandomizeStatements);
		}

		private TranslationResult TryToTranslateReDim(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var reDimStatement = block as ReDimStatement;
			if (reDimStatement == null)
				return null;

			// Any variables that are referenced by the REDIM statement should be treated as if they have been explicitly declared if they haven't
			// otherwise within the current scope. Rather than worry about trying to determine whether or not they have already been declared in
			// the current scope, we'll add them to the returned TranslationResult's ExplicitVariableDeclarations set and them rely upon this
			// being de-duped when it's required. (Note: If this is a ReDim of the function or property return value, where applicable, then
			// do NOT add it to the explicit variable declaration data).
			// - Note: It seems strange that a REDIM should be treated as explicitly declaring a variable, since its purpose is to change the
			//   dimensions of an existing array variable, but VBScript treats it as so (when Option Explicit) is enabled, so we will here too
			// - Note: ONLY add them to the scope if they are not already declared, if the ReDim is in a function and the target variable has
			//   been declared in an outer scope (either within a class or in the outer most scope) then the variable must not then redeclared
			//   within the current function. This must apply even if the variable has only been implicitly declared in the outer scope
			//   (meaning that a variable was accessed in the outer most scope within being explicitly declared first; this can't
			//   happen within a class since statements can only appear in classes within functions or properties). Such
			//   "implicitly declared" variables should be in the ScopeAccessInformation already.
			var explicitVariableDeclarationsToRecord = reDimStatement.Variables
				.Select(v => new
				{
					SourceName = v.Name,
					VariableDeclaration = new VariableDeclaration(
						v.Name,
						VariableDeclarationScopeOptions.Private,
						null
					)
				})
				.Where(newVariable => !scopeAccessInformation.IsDeclaredReference(newVariable.SourceName, _nameRewriter))
				.Where(newVariable =>
					(scopeAccessInformation.ScopeDefiningParent == null) ||
					!scopeAccessInformation.ScopeDefiningParent.Name.Content.Equals(newVariable.SourceName.Content, StringComparison.OrdinalIgnoreCase)
				)
				.ToArray(); // Call ToArray to ensure this is evaluated now and NOT afer the scopeAccessInformation change below

			// These variables are now used to extend the current scopeAccessInformation, this is important for the translation below - any
			// variables that had not be declared should be treated as if they had been explicitly declared (this will affect the results of
			// calls to scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired; it will be able to correctly associate any previously-
			// undeclared variables with the local scope or outer most scope - depending upon whether we're within a function / property or
			// not - without updating the scopeAccessInformation, if we were in the outer most scope then the undeclared variables would be
			// identified as an "Environment References" entry, which we don't want).
			scopeAccessInformation = scopeAccessInformation.ExtendVariables(
				explicitVariableDeclarationsToRecord
					.Select(v => new ScopedNameToken(
						v.SourceName.Content,
						v.SourceName.LineIndex,
						scopeAccessInformation.ScopeDefiningParent.Scope
					))
					.ToNonNullImmutableList()
			);

			var translatedReDimStatements = new NonNullImmutableList<TranslatedStatement>();
			foreach (var variable in reDimStatement.Variables)
			{
				var rewrittenVariableName = _nameRewriter.GetMemberAccessTokenName(variable.Name);
				string targetReference;
				bool isKnownIllegalAssignment;
				if ((scopeAccessInformation.ParentReturnValueNameIfAny != null)
				&& (rewrittenVariableName == _nameRewriter.GetMemberAccessTokenName(scopeAccessInformation.ScopeDefiningParent.Name)))
				{
					// REDIM statements can target the return value of a function - eg.
					//
					//   FUNCTION F1()
					//    REDIM F1(0)
					//    F1(0) = "Result"
					//   END FUNCTION
					//
					// In this case, the "targetReference" value needs to be the name of the temporary value used as the function return
					// value (the ScopeAccessInformation class will always have a non-null references for the ScopeDefiningParent and
					// should have one for the and ParentReturnValueNameIfAny if the ScopeLocation is WithinFunctionOrProperty).
					targetReference = scopeAccessInformation.ParentReturnValueNameIfAny.Name;
					isKnownIllegalAssignment = false;
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
					var targetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(variable.Name, _envRefName, _outerRefName, _nameRewriter);
					targetReference = (targetContainer == null)
						? rewrittenVariableName
						: (targetContainer.Name + "." + rewrittenVariableName);

					// Note: If the target is a CONST then an "Illegal assignment" runtime error must be raised after any evaluation of
					// REDIM arguments has been processed
					var targetReferenceDetails = scopeAccessInformation.TryToGetDeclaredReferenceDetails(variable.Name, _nameRewriter);
					isKnownIllegalAssignment = (targetReferenceDetails != null) && (targetReferenceDetails.ReferenceType == ReferenceTypeOptions.Constant);
				}

				var mayRequireErrorWrapping = scopeAccessInformation.MayRequireErrorWrapping(block);
				if (mayRequireErrorWrapping)
				{
					translatedReDimStatements = translatedReDimStatements.Add(
						new TranslatedStatement(
							GetHandleErrorContent(scopeAccessInformation.ErrorRegistrationTokenIfAny),
							indentationDepth,
							variable.Name.LineIndex
						)
					);
				}

				// If we know that this is an illegal assignment then we need to allow the REDIM's arguments to be evaluated but for an
				// error to be raised after the evaluation. In this case we will not generate code that actually tries to update the
				// target and we will never both trying to validate that the target is an acceptable array even if REDIM PRESERVE
				// is used (since the illegal-assignment error takes precedence)
				// - Note: It is for similar reasons that the original-array-value is the last argument in a RESIZEARRAY call, since
				//   the new dimension sizes must be evaluated before a resize is attempted (meaning that the dimension sizes must
				//   be evaluated before the original array is checked for validity)
				var translatedContentFormat = (reDimStatement.Preserve && !isKnownIllegalAssignment)
					? "{1}.RESIZEARRAY({0}, new object[] {{ {2} }});"
					: "{1}.NEWARRAY(new object[] {{ {2} }});";
				if (!isKnownIllegalAssignment)
					translatedContentFormat = "{0} = " + translatedContentFormat;

				var translatedArguments = new List<string>();
				foreach (var dimension in variable.Dimensions)
				{
					var translatedArgumentDetails = _statementTranslator.Translate(
						dimension,
						scopeAccessInformation,
						ExpressionReturnTypeOptions.Value,
						_logger.Warning
					);
					translatedArguments.Add(translatedArgumentDetails.TranslatedContent);
					var undeclaredVariables = translatedArgumentDetails.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter);
					foreach (var undeclaredVariable in undeclaredVariables)
						_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
					translationResult = translationResult.AddUndeclaredVariables(undeclaredVariables);
				}
				translatedReDimStatements = translatedReDimStatements.Add(
					new TranslatedStatement(
						string.Format(
							translatedContentFormat,
							targetReference,
							_supportRefName.Name,
							string.Join(", ", translatedArguments)
						),
						indentationDepth + (mayRequireErrorWrapping ? 1 : 0),
						variable.Name.LineIndex
					)
				);

				if (mayRequireErrorWrapping)
				{
					translatedReDimStatements = translatedReDimStatements.Add(
						new TranslatedStatement(
							"});",
							indentationDepth,
							variable.Name.LineIndex
						)
					);
				}
				if (isKnownIllegalAssignment)
				{
					translatedReDimStatements = translatedReDimStatements.Add(
						new TranslatedStatement(
							string.Format("_.RAISEERROR(new IllegalAssignmentException({0}));", ("'" + variable.Name.Content + "'").ToLiteral()),
							indentationDepth,
							variable.Name.LineIndex
						)
					);
				}
			}

			return translationResult
				.AddExplicitVariableDeclarations(explicitVariableDeclarationsToRecord.Select(v => v.VariableDeclaration))
				.Add(translatedReDimStatements);
		}

		private TranslationResult TryToTranslateSelect(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var selectBlock = block as SelectBlock;
			if (selectBlock == null)
				return null;

			var codeBlockTranslator = new SelectBlockTranslator(
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
					selectBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		private TranslationResult TryToTranslateStatementOrExpression(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			// This covers Statement and Expression instances as Expression inherits from Statement
			var statementBlock = block as Statement;
			if (statementBlock == null)
				return null;

			var byRefArgumentIdentifier = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
			var byRefArgumentsToRewrite = byRefArgumentIdentifier.GetByRefArgumentsThatNeedRewriting(
				statementBlock.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.None, _logger.Warning),
				scopeAccessInformation,
				new NonNullImmutableList<FuncByRefMapping>()
			);
			int distanceToIdentEvaluationCodeDueToByRefMappings;
			if (byRefArgumentsToRewrite.Any())
			{
				var byRefMappingOpeningTranslationDetails = byRefArgumentsToRewrite.OpenByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
				translationResult = byRefMappingOpeningTranslationDetails.TranslationResult;
				distanceToIdentEvaluationCodeDueToByRefMappings = byRefMappingOpeningTranslationDetails.DistanceToIndentCodeWithMappedValues;
				indentationDepth += distanceToIdentEvaluationCodeDueToByRefMappings;
				scopeAccessInformation = scopeAccessInformation.ExtendVariables(
					byRefArgumentsToRewrite
						.Select(r => new ScopedNameToken(r.To.Name, r.From.LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith))
						.ToNonNullImmutableList()
				);
				statementBlock = byRefArgumentsToRewrite.RewriteStatementUsingByRefArgumentMappings(statementBlock, _nameRewriter);
			}
			else
				distanceToIdentEvaluationCodeDueToByRefMappings = 0;

			var translatedStatementContentDetails = _statementTranslator.Translate(statementBlock, scopeAccessInformation, _logger.Warning);
			var undeclaredVariables = translatedStatementContentDetails.VariablesAccessed
				.Where(v => !scopeAccessInformation.IsDeclaredReference(v, _nameRewriter));
			foreach (var undeclaredVariable in undeclaredVariables)
				_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

			var coreContent = translatedStatementContentDetails.TranslatedContent + ";";
			if (!scopeAccessInformation.MayRequireErrorWrapping(block))
				translationResult = translationResult.Add(new TranslatedStatement(coreContent, indentationDepth, statementBlock.Tokens.First().LineIndex));
			else
			{
				var lineIndexForClosingErrorWrappingContent = statementBlock.Tokens.Last().LineIndex;
				translationResult = translationResult
					.Add(new TranslatedStatement(GetHandleErrorContent(scopeAccessInformation.ErrorRegistrationTokenIfAny), indentationDepth, lineIndexForClosingErrorWrappingContent))
					.Add(new TranslatedStatement(coreContent, indentationDepth + 1, lineIndexForClosingErrorWrappingContent))
					.Add(new TranslatedStatement("});", indentationDepth, lineIndexForClosingErrorWrappingContent));
			}

			if (byRefArgumentsToRewrite.Any())
			{
				indentationDepth -= distanceToIdentEvaluationCodeDueToByRefMappings;
				translationResult = byRefArgumentsToRewrite.CloseByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
			}

			return translationResult.AddUndeclaredVariables(undeclaredVariables);
		}

		protected TranslationResult TryToTranslateValueSettingStatement(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var valueSettingStatement = block as ValueSettingStatement;
			if (valueSettingStatement == null)
				return null;

			var byRefArgumentIdentifier = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
			var byRefArgumentsToRewrite = byRefArgumentIdentifier.GetByRefArgumentsThatNeedRewriting(
				valueSettingStatement.ValueToSet.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning), // TODO: Explain "NotSpecified" (so that ToStageTwoParserExpression doesn't apply the "None" bracket-standardising)
				scopeAccessInformation,
				new NonNullImmutableList<FuncByRefMapping>()
			);
			byRefArgumentsToRewrite = byRefArgumentIdentifier.GetByRefArgumentsThatNeedRewriting(
				valueSettingStatement.Expression.ToStageTwoParserExpression(
					scopeAccessInformation,
					(valueSettingStatement.ValueSetType == ValueSettingStatement.ValueSetTypeOptions.Set) ? ExpressionReturnTypeOptions.Reference : ExpressionReturnTypeOptions.Value,
					_logger.Warning
				),
				scopeAccessInformation,
				byRefArgumentsToRewrite
			);
			int distanceToIdentEvaluationCodeDueToByRefMappings;
			if (byRefArgumentsToRewrite.Any())
			{
				var byRefMappingOpeningTranslationDetails = byRefArgumentsToRewrite.OpenByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
				translationResult = byRefMappingOpeningTranslationDetails.TranslationResult;
				distanceToIdentEvaluationCodeDueToByRefMappings = byRefMappingOpeningTranslationDetails.DistanceToIndentCodeWithMappedValues;
				indentationDepth += distanceToIdentEvaluationCodeDueToByRefMappings;
				scopeAccessInformation = scopeAccessInformation.ExtendVariables(
					byRefArgumentsToRewrite
						.Select(r => new ScopedNameToken(r.To.Name, r.From.LineIndex, ScopeLocationOptions.WithinFunctionOrPropertyOrWith))
						.ToNonNullImmutableList()
				);
				valueSettingStatement = new ValueSettingStatement(
					byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(valueSettingStatement.ValueToSet, _nameRewriter),
					byRefArgumentsToRewrite.RewriteExpressionUsingByRefArgumentMappings(valueSettingStatement.Expression, _nameRewriter),
					valueSettingStatement.ValueSetType
				);
			}
			else
				distanceToIdentEvaluationCodeDueToByRefMappings = 0;

			var translatedValueSettingStatementContentDetails = _valueSettingStatementTranslator.Translate(valueSettingStatement, scopeAccessInformation);
			var undeclaredVariables = translatedValueSettingStatementContentDetails.VariablesAccessed
				.Where(v => !scopeAccessInformation.IsDeclaredReference(v, _nameRewriter));
			foreach (var undeclaredVariable in undeclaredVariables)
				_logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

			var coreContent = translatedValueSettingStatementContentDetails.TranslatedContent + ";";
			if (!scopeAccessInformation.MayRequireErrorWrapping(block))
				translationResult = translationResult.Add(new TranslatedStatement(coreContent, indentationDepth, valueSettingStatement.ValueToSet.Tokens.First().LineIndex));
			else
			{
				var lineIndexForClosingErrorWrappingContent = valueSettingStatement.Expression.Tokens.Last().LineIndex;
				translationResult = translationResult
					.Add(new TranslatedStatement(GetHandleErrorContent(scopeAccessInformation.ErrorRegistrationTokenIfAny), indentationDepth, lineIndexForClosingErrorWrappingContent))
					.Add(new TranslatedStatement(coreContent, indentationDepth + 1, lineIndexForClosingErrorWrappingContent))
					.Add(new TranslatedStatement("});", indentationDepth, lineIndexForClosingErrorWrappingContent));
			}

			if (byRefArgumentsToRewrite.Any())
			{
				indentationDepth -= distanceToIdentEvaluationCodeDueToByRefMappings;
				translationResult = byRefArgumentsToRewrite.CloseByRefReplacementDefinitionWork(translationResult, indentationDepth, _nameRewriter);
			}

			return translationResult.AddUndeclaredVariables(undeclaredVariables);
		}

		private TranslationResult TryToTranslateWith(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var withBlock = block as WithBlock;
			if (withBlock == null)
				return null;

			var codeBlockTranslator = new WithBlockTranslator(
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
					withBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
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
					(scopeLocation == ScopeLocationOptions.WithinFunctionOrPropertyOrWith) ? "object " : "",
					rewrittenName
				);
			}
			else if (!variableDeclaration.ConstantDimensionsIfAny.Any())
			{
				return string.Format(
					"{0}{1} = (object[])null;",
					(scopeLocation == ScopeLocationOptions.WithinFunctionOrPropertyOrWith) ? "object " : "",
					rewrittenName
				);
			}
			return string.Format(
				"{0}{1} = new object[{2}];",
				(scopeLocation == ScopeLocationOptions.WithinFunctionOrPropertyOrWith) ? "object " : "",
				rewrittenName,
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

			// Note: Any repeated variable declarations are ignored - this makes the ReDim translation process easier (where ReDim statements may target
			// already-declared variables or they may be considered to implicitly declare them) but it means that the Dim translation has to do some extra
			// work to pick up on "Name redefined" scenarios. This duplicate removal process will prefer any times that a variable was declared as an array,
			// over when not this is so that if the following was present:
			//   DIM a()
			//   REDIM a(12)
			// Then it would definitely result in there only being a single explicit variable declaration - "object a = (object[])null;" - the variable
			// declaration arising from the REDIM must be ignored.
			var uniqueExplicitVariableDeclarations = translationResult.ExplicitVariableDeclarations
				.Select((v, i) => new { Index = i, Variable = v })
				.GroupBy(indexedVariable => indexedVariable.Variable.Name.Content)
				.Select(group => group.OrderBy(v => (v.Variable.ConstantDimensionsIfAny == null) ? 1 : 0))
				.Select(group => group.First())
				.OrderBy(indexedVariable => indexedVariable.Index)
				.Select(indexedVariable => indexedVariable.Variable)
				.ToNonNullImmutableList();

			var variableDeclarationStatements = new NonNullImmutableList<TranslatedStatement>();
			foreach (var explicitVariableDeclaration in uniqueExplicitVariableDeclarations)
			{
				variableDeclarationStatements = variableDeclarationStatements.Add(new TranslatedStatement(
					TranslateVariableInitialisation(explicitVariableDeclaration, ScopeLocationOptions.WithinFunctionOrPropertyOrWith),
					indentationDepthForExplicitVariableDeclarations,
					explicitVariableDeclaration.Name.LineIndex
				));
			}
			return new TranslationResult(
				variableDeclarationStatements.AddRange(translationResult.TranslatedStatements),
				new NonNullImmutableList<VariableDeclaration>(),
				translationResult.UndeclaredVariablesAccessed
			);
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
				// If this variable was not explicitly declared with a DIM, but WAS referenced by a REDIM, then VBScript will not consider
				// it undeclared - in this case, there will be an ExplicitVariableDeclarations entry from the REDIM which must be checked
				// for here before declaring the variable officially undeclared.
				if (translationResult.ExplicitVariableDeclarations.Any(v => _nameRewriter.AreNamesEquivalent(v.Name, undeclaredVariable)))
					continue;

				var rewrittenName = _nameRewriter.GetMemberAccessTokenName(undeclaredVariable);
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
							TranslateVariableInitialisation(v, ScopeLocationOptions.WithinFunctionOrPropertyOrWith) + " /* Undeclared in source */",
							indentationDepthForExplicitVariableDeclarations,
							v.Name.LineIndex
						)
					)
					.ToNonNullImmutableList()
					.AddRange(translationResult.TranslatedStatements),
				translationResult.ExplicitVariableDeclarations,
				new NonNullImmutableList<NameToken>()
			);
		}

		private string GetHandleErrorContent(CSharpName errorRegistrationToken)
		{
			if (errorRegistrationToken == null)
				throw new ArgumentNullException("errorRegistrationToken");

			return string.Format(
				"{0}.HANDLEERROR({1}, () => {{",
				_supportRefName.Name,
				errorRegistrationToken.Name
			);
		}

		/// <summary>
		/// These are the block translators which apply within functions or properties. They also apply to the outer most scope, but so do some other (such as
		/// TryToTranslateClass and TryToTranslateFunction). These won't apply to all scopes - for example, the ClassBlockTranslator can't accept statements
		/// that should appear within functions - such as IF blocks.
		/// </summary>
		protected NonNullImmutableList<BlockTranslationAttempter>  GetWithinFunctionBlockTranslators()
		{
			return new NonNullImmutableList<BlockTranslationAttempter>(
				new BlockTranslationAttempter[]
				{
					this.TryToTranslateBlankLine,
					this.TryToTranslateComment,
					this.TryToTranslateConst,
					this.TryToTranslateDim,
					this.TryToTranslateDo,
					this.TryToTranslateExit,
					this.TryToTranslateErase,
					this.TryToTranslateFor,
					this.TryToTranslateForEach,
					this.TryToTranslateIf,
					this.TryToTranslateOnErrorResumeNext,
					this.TryToTranslateOnErrorGotoZero,
					this.TryToTranslateReDim,
					this.TryToTranslateRandomize,
					this.TryToTranslateStatementOrExpression,
					this.TryToTranslateSelect,
					this.TryToTranslateValueSettingStatement,
					this.TryToTranslateWith
				}
			);
		}
	}
}
