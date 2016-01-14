using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation
{
	public interface ITranslateValueSettingsStatements
    {
		/// <summary>
		/// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
		/// </summary>
		TranslatedStatementContentDetails Translate(ValueSettingStatement statement, ScopeAccessInformation scopeAccessInformation);
    }
}
