using CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class ScopeAccessInformation_Extensions
    {
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
			IHaveNestedContent parentIfAny,
			IDefineScope scopeDefiningParentIfAny,
			IEnumerable<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            blocks = blocks.ToNonNullImmutableList().FlattenAllAccessibleBlockLevelCodeBlocks();
            return new ScopeAccessInformation(
				parentIfAny,
				scopeDefiningParentIfAny,
                scopeInformation.Classes.AddRange(
                    blocks
                        .Where(b => b is ClassBlock)
                        .Cast<ClassBlock>()
                        .Select(c => c.Name)
                ),
                scopeInformation.Functions.AddRange(
                    blocks
                        .Where(b => (b is FunctionBlock) || (b is SubBlock))
                        .Cast<AbstractFunctionBlock>()
                        .Select(f => f.Name)
                ),
                scopeInformation.Properties.AddRange(
                    blocks
                        .Where(b => b is PropertyBlock)
                        .Cast<PropertyBlock>()
                        .Select(p => p.Name)
                ),
                scopeInformation.Variables.AddRange(
                    blocks
                        .Where(b => b is DimStatement) // This covers DIM, REDIM, PRIVATE and PUBLIC (they may all be considered the same for these purposes)
                        .Cast<DimStatement>()
                        .SelectMany(d => d.Variables.Select(v => v.Name))
                )
            );
        }

		/// <summary>
		/// If the parentIfAny is scope-defining then both the parentIfAny and scopeDefiningParentIfAny references will be set to it, this is convenience
		/// method to save having to specify it explicitly for both
		/// </summary>
		public static ScopeAccessInformation Extend(this ScopeAccessInformation scopeInformation, IDefineScope parentIfAny, IEnumerable<ICodeBlock> blocks)
		{
			if (scopeInformation == null)
				throw new ArgumentNullException("scopeInformation");
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			return Extend(scopeInformation, parentIfAny, parentIfAny, blocks);
		}
	}
}
