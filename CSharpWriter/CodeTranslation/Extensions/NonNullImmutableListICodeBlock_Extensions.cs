using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class NonNullImmutableListICodeBlock_Extensions
    {
        // TODO: Pull this into ScopeAccessInformation_Extensions if it doesn't end up getting used anywhere else
        public static NonNullImmutableList<ICodeBlock> FlattenAllAccessibleBlockLevelCodeBlocks(this NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            var flattenedBlocks = new NonNullImmutableList<ICodeBlock>();
            foreach (var block in blocks)
            {
                flattenedBlocks = flattenedBlocks.Add(block);

                var parentBlock = block as IHaveNestedContent;
                if (parentBlock == null)
                    continue;

                if (parentBlock is IDefineScope)
                {
                    // If this defines scope then we can't expand the current scope by drilling into it - eg. if the current block
                    // is a class then it has nested statements but we can't access them directly (we can't call a function on a
                    // class without calling it on an instance of that class)
                    continue;
                }

                flattenedBlocks = flattenedBlocks.AddRange(
                    FlattenAllAccessibleBlockLevelCodeBlocks(
                        parentBlock.AllExecutableBlocks.ToNonNullImmutableList()
                    )
                );
            }
            return flattenedBlocks;
        }
    }
}
