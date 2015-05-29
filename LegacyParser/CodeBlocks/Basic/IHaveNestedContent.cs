using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// This represents a code block that is built up of other code blocks. The FunctionBlock contains other code blocks and acts as a "scope-defining"
    /// parent (any undeclared variables within in are implicitly declared within its scope, not the outer scope around it). The IfBlock contains other
    /// blocks for the conditions and for the statements that are executed when those conditions are met (though it does act as a scope-defining parent).
    /// </summary>
    public interface IHaveNestedContent : ICodeBlock
    {
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        IEnumerable<ICodeBlock> AllExecutableBlocks { get; }
    }
}
