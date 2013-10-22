using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSharpSupport;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class CodeBlockTranslator
    {
        private readonly CSharpName _supportClassName;
        private readonly VBScriptNameRewriter _nameRewriter;
        private readonly TempValueNameGenerator _tempNameGenerator;
        private readonly ITranslateIndividualStatements _statementTranslator;
        public CodeBlockTranslator(
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

        public NonNullImmutableList<TranslatedStatement> Translate(NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            var translationResult = Translate(
                blocks,
                ScopeAccessInformation.Empty.Extend(ParentConstructTypeOptions.None, blocks),
                0
            );
            translationResult = FlushExplicitVariableDeclarations(translationResult, ParentConstructTypeOptions.None, 0);
            translationResult = FlushUndeclaredVariableDeclarations(translationResult, 0);
            return translationResult.TranslatedStatements;
        }

        private TranslationResult Translate(
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeAccessInformation scopeAccessInformation,
            int indentationDepth)
        {
            if (blocks == null)
                throw new ArgumentNullException("block");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var translationResult = TranslationResult.Empty;
			foreach (var block in RemoveDuplicateFunctions(blocks))
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
                    translationResult = translationResult.Add(
                        TranslateClassHeader(classBlock, indentationDepth),
                        Translate(
                            classBlock.Statements.ToNonNullImmutableList(),
                            scopeAccessInformation.Extend(
                                ParentConstructTypeOptions.Class,
                                classBlock.Statements
                            ),
                            indentationDepth + 1
                        ),
                        new TranslatedStatement("}", indentationDepth)
                    );
                    continue;
                }

                var functionBlock = ((block is FunctionBlock) || (block is SubBlock)) ? block as AbstractFunctionBlock : null;
                if (functionBlock != null)
                {
                    translationResult = translationResult.Add(
                        TranslateFunctionHeader(
                            functionBlock,
                            (block is FunctionBlock), // hasReturnValue (true for FunctionBlock, false for SubBlock)
                            indentationDepth
                        ),
                        Translate(
                            functionBlock.Statements.ToNonNullImmutableList(),
                            scopeAccessInformation.Extend(
                                ParentConstructTypeOptions.FunctionOrProperty,
                                functionBlock.Statements
                            ),
                            indentationDepth + 1
                        ),
                        new TranslatedStatement("}", indentationDepth)
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
                    throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
                }

                throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");

                // TODO
                // - DoBlock
                // - ExitStatement
                // - Expression / Statement / ValueSettingStatement
                // - ForBlock
                // - ForEachBlock
                // - OnErrorResumeNext / OnErrorGoto0
                // - PropertyBlock (check for Default on the Get - the only place it's valid - before rendering Let or Set)
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

		/// <summary>
		/// VBScript allows functions with the same name to appear multiple times, where all but the last implementation will be ignored. This is not
		/// allowed within classes, but this translation should only be dealing with valid VBScript so there will be no validation for that here.
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

				var functionName = _nameRewriter(functionBlock.Name).Name;
				removeAtLocations.AddRange(
					blocks
						.Select((b, blockIndex) => new { Index = blockIndex, Block = b })
						.Where(indexedBlock => indexedBlock.Block is AbstractFunctionBlock)
						.Where(indexedBlock => _nameRewriter(((AbstractFunctionBlock)indexedBlock.Block).Name).Name == functionName)
						.Select(indexedBlock => indexedBlock.Index)
						.OrderByDescending(blockIndex => blockIndex)
						.Skip(1) // Leave the last one intact
				);
			}
			foreach (var removeIndex in removeAtLocations.Distinct().OrderByDescending(i => i))
				blocks = blocks.RemoveAt(removeIndex);
			return blocks;
		}

		private string TranslateVariableDeclaration(VariableDeclaration variableDeclaration)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");

            return string.Format(
                "object {0} = {1}null;",
                _nameRewriter(variableDeclaration.Name).Name,
                variableDeclaration.IsArray ? "(object[])" : ""
            );
        }

        private IEnumerable<TranslatedStatement> TranslateClassHeader(ClassBlock classBlock, int indentationDepth)
        {
            if (classBlock == null)
                throw new ArgumentNullException("classBlock");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var className = _nameRewriter.GetMemberAccessTokenName(classBlock.Name);
            return new[]
            {
                new TranslatedStatement("public class " + className, indentationDepth),
                new TranslatedStatement("{", indentationDepth),
                new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionality).FullName + " " + _supportClassName.Name + ";", indentationDepth + 1),
                new TranslatedStatement("public " + className + "(" + typeof(IProvideVBScriptCompatFunctionality).FullName + " compatLayer)", indentationDepth + 1),
                new TranslatedStatement("{", indentationDepth + 1),
                new TranslatedStatement("if (compatLayer == null)", indentationDepth + 2),
                new TranslatedStatement("throw new ArgumentNullException(compatLayer)", indentationDepth + 3),
                new TranslatedStatement("this." + _supportClassName.Name + " = compatLayer;", indentationDepth + 2),
                new TranslatedStatement("}", indentationDepth + 1)
            };
        }

        private IEnumerable<TranslatedStatement> TranslateFunctionHeader(AbstractFunctionBlock functionBlock, bool hasReturnValue, int indentationDepth)
        {
            if (functionBlock == null)
                throw new ArgumentNullException("functionBlock");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var content = new StringBuilder();
            content.Append(functionBlock.IsPublic ? "public" : "private");
            content.Append(" ");
            content.Append(hasReturnValue ? "object" : "void");
            content.Append(" ");
            content.Append(_nameRewriter.GetMemberAccessTokenName(functionBlock.Name));
            content.Append("(");
            var numberOfParameters = functionBlock.Parameters.Count();
            for (var index = 0; index < numberOfParameters; index++)
            {
                var parameter = functionBlock.Parameters.ElementAt(index);
                if (parameter.ByRef)
                    content.Append("ref ");
                content.Append(_nameRewriter.GetMemberAccessTokenName(parameter.Name));
                if (index < (numberOfParameters - 1))
                    content.Append(", ");
            }
            content.Append(")");

			var translatedStatements = new List<TranslatedStatement>();
			if (functionBlock.IsDefault)
				translatedStatements.Add(new TranslatedStatement("[" + typeof(IsDefault).FullName + "]", indentationDepth));
			translatedStatements.Add(new TranslatedStatement(content.ToString(), indentationDepth));
			translatedStatements.Add(new TranslatedStatement("{", indentationDepth));
			return translatedStatements;
        }

		private TranslationResult FlushExplicitVariableDeclarations(
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

		/// <summary>
        /// This should only performed at the outer layer (and so no ParentConstructTypeOptions value is required, it is assumed to be None)
        /// </summary>
        private TranslationResult FlushUndeclaredVariableDeclarations(TranslationResult translationResult, int indentationDepth)
        {
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            return new TranslationResult(
                translationResult.UndeclaredVariablesAccessed
                    .Select(v =>
                         new TranslatedStatement(
                            TranslateVariableDeclaration(
                                // Undeclared variables will be specified as non-array types initially (hence the false
                                // value for the isArray argument if the VariableDeclaration constructor call below)
                                new VariableDeclaration(v, VariableDeclarationScopeOptions.Public, false)
                            ),
                            indentationDepth
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
