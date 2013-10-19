using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    public interface IHaveNestedContent : ICodeBlock
    {
        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> AllExecutableBlocks { get; }
    }
}
