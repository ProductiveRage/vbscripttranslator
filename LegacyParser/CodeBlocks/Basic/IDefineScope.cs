namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// This interface is applied to code blocks with nested statements that define block-level scopes - eg. classes, functions and properties
    /// but not if blocks or while loops. Note that this also applies to VBScript error-trapping; an "On Error Resume Next" within a for loop
    /// will be maintained after the loop has completed (and so the ForBlock does not implement this interface) but will be cleared once a
    /// function ends (and so the FunctionBlock does implement this).
    /// </summary>
    public interface IDefineScope : IHaveNestedContent { }
}
