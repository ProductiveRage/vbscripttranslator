namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// This is a marker interface to differentiate between constructs that have child code blocks that are expected to be executed once and those
    /// that will loop over their content
    /// </summary>
    public interface ILoopOverNestedContent : IHaveNestedContent { }
}
