using VBScriptTranslator.CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions
{
    public static class IEnumerableOfCodeBlocks_Extensions
    {
        public static bool DoesScopeContainOnErrorResumeNext(this IEnumerable<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            foreach (var block in blocks)
            {
                if (block == null)
                    throw new ArgumentException("Null reference encountered in blocks set");

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
            }
            return false;
        }
    }
}
