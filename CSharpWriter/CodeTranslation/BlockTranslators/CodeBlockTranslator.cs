using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public abstract class CodeBlockTranslator
    {
		protected readonly CSharpName _supportClassName;
		protected readonly VBScriptNameRewriter _nameRewriter;
		protected readonly TempValueNameGenerator _tempNameGenerator;
		private readonly ITranslateIndividualStatements _statementTranslator;
		private readonly ITranslateValueSettingsStatements _valueSettingStatementTranslator;
        private readonly ILogInformation _logger;
        protected CodeBlockTranslator(
            CSharpName supportClassName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator,
			ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
        {
            if (supportClassName == null)
                throw new ArgumentNullException("supportClassName");
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

            _supportClassName = supportClassName;
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

                    translationResult = blockTranslationResult;
					hasBlockBeenTranslated = true;
                    break;
				}
				if (!hasBlockBeenTranslated)
					throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
            }

            // If the current parent construct doesn't affect scope (like IF and WHILE and unlike CLASS and FUNCTION) then the translationResult
            // can be returned directly and the nearest construct that does affect scope will be responsible for translating any explicit
            // variable declarations into translated statements
			if (!(scopeAccessInformation.ParentIfAny is IDefineScope))
                return translationResult;
            
            return FlushExplicitVariableDeclarations(
                translationResult,
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
                _supportClassName,
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
			// This covers the DimStatement, ReDimStatement, PrivateVariableStatement and PublicVariableStatement
			var explicitVariableDeclarationBlock = block as DimStatement;
			if (explicitVariableDeclarationBlock == null)
				return null;

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

			var areDimensionsRequired = (
				(explicitVariableDeclarationBlock is ReDimStatement) ||
				(explicitVariableDeclarationBlock.Variables.Any(v => (v.Dimensions != null) && (v.Dimensions.Count > 0)))
			);
			if (!areDimensionsRequired)
				return translationResult;

			// TODO: If this is a ReDim then non-constant expressions may be used to set the dimension limits, in which case it may not be moved
			// (though a default-null declaration SHOULD be added as well as leaving the ReDim translation where it is)

            // TODO: Is this is a ReDim targetting the name of the function that it's in then an "Illegal assignment" error will be raised when
            // the code is executed (trying to Dim the name of the function will result in a "Name redefined" error as soon as the script is
            // exceuted, not just when the line in question is executed
			
            // TODO: Need a translated statement if setting dimensions
            throw new NotImplementedException("Not enabled support for declaring array variables with specifid dimensions yet");
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
                _supportClassName,
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

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
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

            var translatedStatementContentDetails = _statementTranslator.Translate(statementBlock, scopeAccessInformation);
			return
                translationResult.Add(
				    new TranslatedStatement(
                        translatedStatementContentDetails.TranslatedContent + ";",
					    indentationDepth
				    )
			    ).Add(
                    scopeAccessInformation.GetUndeclaredVariables(
                        translatedStatementContentDetails.VariablesAccesed
                    )
                );
		}

		protected TranslationResult TryToTranslateValueSettingStatement(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var valueSettingStatement = block as ValueSettingStatement;
			if (valueSettingStatement == null)
				return null;

			var translatedValueSettingStatementContentDetails = _valueSettingStatementTranslator.Translate(valueSettingStatement, scopeAccessInformation);
			return
				translationResult.Add(
					new TranslatedStatement(
						translatedValueSettingStatementContentDetails.TranslatedContent + ";",
						indentationDepth
					)
				).Add(
					scopeAccessInformation.GetUndeclaredVariables(
						translatedValueSettingStatementContentDetails.VariablesAccesed
					)
				);
		}

        protected string TranslateVariableDeclaration(VariableDeclaration variableDeclaration)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");

            return string.Format(
                "object {0} = {1}null;",
                _nameRewriter.GetMemberAccessTokenName(variableDeclaration.Name),
                variableDeclaration.IsArray ? "(object[])" : ""
            );
        }

		protected TranslationResult FlushExplicitVariableDeclarations(TranslationResult translationResult, int indentationDepth)
		{
			// TODO: Consider trying to insert the content after any comments or blank lines?
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			return new TranslationResult(
				translationResult.ExplicitVariableDeclarations
					.Select(v =>
						 new TranslatedStatement(TranslateVariableDeclaration(v), indentationDepth)
					)
					.ToNonNullImmutableList()
					.AddRange(translationResult.TranslatedStatements),
				new NonNullImmutableList<VariableDeclaration>(),
				translationResult.UndeclaredVariablesAccessed
			);
		}
    }
}
