using LegacyParser = VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using StageTwoParser = VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation
{
    public interface ITranslateIndividualStatements
    {
		/// <summary>
		/// This will never return null or blank, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
		/// </summary>
        string Translate(LegacyParser.Statement statement, ScopeAccessInformation scopeAccessInformation);
        
		/// <summary>
        /// This will never return null or blank, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
		/// </summary>
        string Translate(LegacyParser.Expression expression, ScopeAccessInformation scopeAccessInformation, ExpressionReturnTypeOptions returnRequirements);

        /// <summary>
        /// This will never return null or blank, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
        /// </summary>
        string Translate(StageTwoParser.Expression expression, ScopeAccessInformation scopeAccessInformation, ExpressionReturnTypeOptions returnRequirements);
    }
}
