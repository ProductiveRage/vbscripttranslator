namespace VBScriptTranslator.LegacyParser.CodeBlocks
{
    public interface ICodeBlock
    {
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter);
    }
}
