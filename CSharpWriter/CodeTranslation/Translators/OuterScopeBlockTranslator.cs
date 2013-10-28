using System;
using System.Collections.Generic;
using System.Linq;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class OuterScopeBlockTranslator : CodeBlockTranslator
    {
		public OuterScopeBlockTranslator(
            CSharpName supportClassName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator) : base(supportClassName, nameRewriter, tempNameGenerator, statementTranslator) { }

        public NonNullImmutableList<TranslatedStatement> Translate(NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

			blocks = RemoveDuplicateFunctions(blocks);
			var translationResult = Translate(
                blocks,
                ScopeAccessInformation.Empty.Extend(null, blocks),
                0
            );
            translationResult = FlushExplicitVariableDeclarations(translationResult, 0);
            translationResult = FlushUndeclaredVariableDeclarations(translationResult, 0);
            return translationResult.TranslatedStatements;
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
				new BlockTranslationAttempter[]
				{
					base.TryToTranslateBlankLine,
					base.TryToTranslateClass,
					base.TryToTranslateComment,
					base.TryToTranslateDim,
					base.TryToTranslateDo,
					base.TryToTranslateExit,
					base.TryToTranslateFor,
					base.TryToTranslateForEach,
					base.TryToTranslateFunction,
					base.TryToTranslateIf,
					base.TryToTranslateOptionExplicit,
					base.TryToTranslateRandomize,
					base.TryToTranslateStatementOrExpression,
					base.TryToTranslateSelect,
                    base.TryToTranslateValueSettingStatement
				}.ToNonNullImmutableList(),
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

				var functionName = _nameRewriter(functionBlock.Name).Name;
				removeAtLocations.AddRange(
					blocks
						.Select((b, blockIndex) => new { Index = blockIndex, Block = b })
						.Where(indexedBlock => indexedBlock.Block is AbstractFunctionBlock)
						.Where(indexedBlock => _nameRewriter(((AbstractFunctionBlock)indexedBlock.Block).Name).Name == functionName)
						.Select(indexedBlock => indexedBlock.Index)
						.OrderByDescending(blockIndex => blockIndex).Skip(1) // Leave the last one intact
				);
			}
			foreach (var removeIndex in removeAtLocations.Distinct().OrderByDescending(i => i))
				blocks = blocks.RemoveAt(removeIndex);
			return blocks;
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
							base.TranslateVariableDeclaration(
								// Undeclared variables will be specified as non-array types initially (hence the false
								// value for the isArray argument if the VariableDeclaration constructor call below)
								new VariableDeclaration(v, VariableDeclarationScopeOptions.Public, false)
							),
							indentationDepth
						)
					)
                    .GroupBy(s => s.Content).Select(group => group.First()) // Lazy way to do distinct
                    .OrderBy(s => s.Content)
					.ToNonNullImmutableList()
                    .Add(new TranslatedStatement("", indentationDepth)) // Blank line between inject variable declarations and the rest of the generated code
					.AddRange(translationResult.TranslatedStatements),
				translationResult.ExplicitVariableDeclarations,
				new NonNullImmutableList<NameToken>()
			);
		}
	}
}
