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
    public class OuterScopeBlockTranslator : CodeBlockTranslator
    {
		public OuterScopeBlockTranslator(
            CSharpName supportRefName,
            CSharpName envRefName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator,
            ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
            : base(supportRefName, envRefName, outerRefName, nameRewriter, tempNameGenerator, statementTranslator, valueSettingStatementTranslator, logger) { }

        public NonNullImmutableList<TranslatedStatement> Translate(NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            // Note: There is no need to check for identically-named classes since that would cause a "Name Redefined" error even if Option Explicit was not enabled
            blocks = RemoveDuplicateFunctions(blocks);

            // TODO: Get all non-class code blocks together (including functions), wrapping it into a static "Program" class / function (allow t to be specified through
            // constructor argument?) with static functions for the function code blocks and an "env" object with all of the required properties specified on it (?)
            var classBlocks = blocks.Where(block => block is ClassBlock);
            var otherBlocks = blocks.Where(block => !(block is ClassBlock));

            //var environmentClassBlock = GetEnvironmentClassBlock(blocks);

			var translationResult = Translate(
                blocks,
                ScopeAccessInformation.Empty.Extend(null, blocks, ScopeLocationOptions.OutermostScope),
                1 // indentationDepth
            );
            translationResult = FlushExplicitVariableDeclarations(translationResult, 0);
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
