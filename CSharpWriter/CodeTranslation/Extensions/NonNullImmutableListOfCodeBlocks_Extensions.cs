using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class NonNullImmutableListOfCodeBlocks_Extensions
    {
        public static bool DoesScopeContainOnErrorResumeNext(this NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            foreach (var block in blocks)
            {
                if (block is OnErrorResumeNext)
                    return true;

                if (block is IDefineScope)
                {
                    // If this block defines its own scope then it doesn't matter if it contains an OnErrorResumeNext statement
                    // within it, it won't affect this scope
                    continue;
                }

                var blockWithNestedContent = block as IHaveNestedContent;
                if ((blockWithNestedContent != null) && blockWithNestedContent.AllExecutableBlocks.ToNonNullImmutableList().DoesScopeContainOnErrorResumeNext())
                    return true;
                    continue;
            }
            return false;
        }
    }
}
