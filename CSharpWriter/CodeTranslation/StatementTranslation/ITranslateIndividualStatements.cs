using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
    public interface ITranslateIndividualStatements
    {
        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
        /// </summary>
        TranslatedStatementContentDetails Translate(Expression expression, ScopeAccessInformation scopeAccessInformation, ExpressionReturnTypeOptions returnRequirements);
    }
}
