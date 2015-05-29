using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// This interface is used for code blocks that directly wrap expressions but that don't inherit from the Statement class - it allows the Seed Expression
    /// to be retrieved from the RandomizeStatement statement, for example, or for the two expressions to be retrieved from the ValueSettingStatement. This
    /// is not appropriate for code blocks that act as parents for other statements, these are represented by the IHaveNestedContent interface.
    /// </summary>
    public interface IHaveNonNestedExpressions : ICodeBlock
    {
        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> NonNestedExpressions { get; }
    }
}
