using System.Collections.Generic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
    public interface ITranslateIndividualStatements
    {
        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
        /// </summary>
        TranslatedStatementContentDetails Translate(Expression expression, ScopeAccessInformation scopeAccessInformation, ExpressionReturnTypeOptions returnRequirements);

        /// <summary>
        /// This generates the content that initialises a new IProvideCallArguments instance, based upon the specified argument values. This will throw
        /// an exception for null arguments or an argumentValues set containing any null references. It will never return null, it will raise an exception
        /// if unable to satisfy the request.
        /// </summary>
        TranslatedStatementContentDetails TranslateAsArgumentProvider(
            IEnumerable<Expression> argumentValues,
            ScopeAccessInformation scopeAccessInformation,
            bool forceAllArgumentsToBeByVal
        );
    }
}
