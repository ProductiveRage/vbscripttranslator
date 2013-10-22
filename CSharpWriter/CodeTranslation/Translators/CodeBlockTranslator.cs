using System;
using System.Collections.Generic;
using System.Linq;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation
{
    public abstract class CodeBlockTranslator
    {
		protected readonly CSharpName _supportClassName;
		protected readonly VBScriptNameRewriter _nameRewriter;
		protected readonly TempValueNameGenerator _tempNameGenerator;
		protected readonly ITranslateIndividualStatements _statementTranslator;
        protected CodeBlockTranslator(
            CSharpName supportClassName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator)
        {
            if (supportClassName == null)
                throw new ArgumentNullException("supportClassName");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (tempNameGenerator == null)
                throw new ArgumentNullException("tempNameGenerator");
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");

            _supportClassName = supportClassName;
            _nameRewriter = nameRewriter;
            _tempNameGenerator = tempNameGenerator;
            _statementTranslator = statementTranslator;
        }

        protected TranslationResult TranslateCommon(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
        {
            if (blocks == null)
                throw new ArgumentNullException("block");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var translationResult = TranslationResult.Empty;
			foreach (var block in blocks)
            {
                if (block is OptionExplicit)
                    continue;

                if (block is BlankLine)
                {
                    translationResult = translationResult.Add(
                        new TranslatedStatement("", indentationDepth)
                    );
                    continue;
                }

                var commentBlock = block as CommentStatement;
                if (commentBlock != null)
                {
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
                            continue;
                        }
                    }
                    translationResult = translationResult.Add(
                        new TranslatedStatement(translatedCommentContent, indentationDepth)
                    );
                    continue;
                }

                // This covers the DimStatement, ReDimStatement, PrivateVariableStatement and PublicVariableStatement
                var explicitVariableDeclarationBlock = block as DimStatement;
                if (explicitVariableDeclarationBlock != null)
                {
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
                        continue;
                    
                    // TODO: If this is a ReDim then non-constant expressions may be used to set the dimension limits, in which case it may not be moved
                    // (though a default-null declaration SHOULD be added as well as leaving the ReDim translation where it is)
                    // TODO: Need a translated statement if setting dimensions
                    throw new NotImplementedException("Not enabled support for declaring array variables with specifid dimensions yet");
                }

                var classBlock = block as ClassBlock;
                if (classBlock != null)
                {
					var codeBlockTranslator = new ClassBlockTranslator(_supportClassName, _nameRewriter, _tempNameGenerator, _statementTranslator);
					translationResult = translationResult.Add(
						codeBlockTranslator.Translate(
							classBlock,
							scopeAccessInformation,
							indentationDepth
						)
					);
					continue;
                }

                var functionBlock = ((block is FunctionBlock) || (block is SubBlock)) ? block as AbstractFunctionBlock : null;
                if (functionBlock != null)
                {
					var codeBlockTranslator = new FunctionBlockTranslator(_supportClassName, _nameRewriter, _tempNameGenerator, _statementTranslator);
					translationResult = translationResult.Add(
						codeBlockTranslator.Translate(
							functionBlock,
							scopeAccessInformation,
							indentationDepth
						)
					);
                    continue;
                }

                // This covers Statement and Expression instances
                var statementBlock = block as Statement;
                if (statementBlock != null)
                {
                    translationResult = translationResult.Add(
                        new TranslatedStatement(
                            _statementTranslator.Translate(statementBlock) + ";",
                            indentationDepth
                        )
                    );
                    continue;
                }

                var valueSettingStatement = block as ValueSettingStatement;
                if (valueSettingStatement != null)
                {
                    // TODO: This isn't right, we need to access the target reference in a manner that allows us to change it.
                    // With a _.SET method?
                    /*
                    translationResult = translationResult.Add(
                        new TranslatedStatement(
                            string.Format(
                                "{0} = {1};",
                                _statementTranslator.Translate(
                                    valueSettingStatement.ValueToSet,
                                    ExpressionReturnTypeOptions.NotSpecified
                                ),
                                _statementTranslator.Translate(
                                    valueSettingStatement.Expression,
                                    (valueSettingStatement.ValueSetType == ValueSettingStatement.ValueSetTypeOptions.Set)
                                        ? ExpressionReturnTypeOptions.Reference
                                        : ExpressionReturnTypeOptions.Value
                                )
                            ),
                            indentationDepth
                        )
                    );
                    continue;
                     */
                    throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
                }

                var ifBlock = block as IfBlock;
				if (ifBlock != null)
				{
					translationResult = translationResult.Add(
						TranslateIfBlock(ifBlock, indentationDepth)
					);
					continue;
				}

                throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");

                // TODO
                // - DoBlock
                // - ExitStatement (only in Function, Property, For and While code block translators)
                // - Expression / Statement / ValueSettingStatement
                // - ForBlock
                // - ForEachBlock
                // - OnErrorResumeNext / OnErrorGoto0
                // - RandomizeStatement (see http://msdn.microsoft.com/en-us/library/e566zd96(v=vs.84).aspx when implementing RND)
                // - SelectBlock

                // Error on particular statements encountered out of context
                // - This is a "RegularProcessor"..? (as opposed to ClassProcessor which allows properties but not many other things)
                // 1. Exit Statement (should always be within another construct - eg. Do, For)
                // 2. Properties
            }

            // If the current parent construct doesn't affect scope (like IF and WHILE and unlike CLASS and FUNCTION) then the translationResult
            // can be returned directly and the nearest construct that does affect scope will be responsible for translating any explicit
            // variable declarations into translated statements
            if (scopeAccessInformation.ParentConstructType == ParentConstructTypeOptions.NonScopeAlteringConstruct)
                return translationResult;
            
            return FlushExplicitVariableDeclarations(
                translationResult,
                scopeAccessInformation.ParentConstructType,
                indentationDepth
            );
        }

		protected string TranslateVariableDeclaration(VariableDeclaration variableDeclaration)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");

            return string.Format(
                "object {0} = {1}null;",
                _nameRewriter(variableDeclaration.Name).Name,
                variableDeclaration.IsArray ? "(object[])" : ""
            );
        }

        private IEnumerable<TranslatedStatement> TranslateIfBlock(IfBlock ifBlock, int indentationDepth)
        {
            if (ifBlock == null)
                throw new ArgumentNullException("functionBlock");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            //var hadFirstClause = false;
            foreach (var clause in ifBlock.Clauses)
            {
            }
            throw new NotImplementedException(); // TODO
        }

		protected TranslationResult FlushExplicitVariableDeclarations(
			TranslationResult translationResult,
			ParentConstructTypeOptions parentConstructType,
			int indentationDepth)
		{
			// TODO: Consider trying to insert the content after any comments or blank lines?
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
			if (!Enum.IsDefined(typeof(ParentConstructTypeOptions), parentConstructType))
				throw new ArgumentOutOfRangeException("parentConstructType");
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
