using VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions
{
    public static class ITranslateIndividualStatements_Extensions
    {
        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
        /// </summary>
        public static TranslatedStatementContentDetails Translate(
            this ITranslateIndividualStatements statementTranslator,
            Statement statement,
            ScopeAccessInformation scopeAccessInformation,
            Action<string> warningLogger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
            if (statement == null)
                throw new ArgumentNullException("statement");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (warningLogger == null)
                throw new ArgumentNullException("warningLogger");

            return Translate(statementTranslator, statement, scopeAccessInformation, ExpressionReturnTypeOptions.None, warningLogger);
        }

        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
        /// </summary>
        public static TranslatedStatementContentDetails Translate(
            this ITranslateIndividualStatements statementTranslator,
            Expression expression,
            ScopeAccessInformation scopeAccessInformation,
            ExpressionReturnTypeOptions returnRequirements,
            Action<string> warningLogger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
                throw new ArgumentOutOfRangeException("returnRequirements");
            if (warningLogger == null)
                throw new ArgumentNullException("warningLogger");

            return Translate(statementTranslator, (Statement)expression, scopeAccessInformation, returnRequirements, warningLogger);
        }

        private static TranslatedStatementContentDetails Translate(
            ITranslateIndividualStatements statementTranslator,
            Statement statement,
            ScopeAccessInformation scopeAccessInformation,
            ExpressionReturnTypeOptions returnRequirements,
            Action<string> warningLogger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
            if (statement == null)
                throw new ArgumentNullException("statement");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
                throw new ArgumentOutOfRangeException("returnRequirements");
            if (warningLogger == null)
                throw new ArgumentNullException("warningLogger");

            return statementTranslator.Translate(
                statement.ToStageTwoParserExpression(scopeAccessInformation, returnRequirements, warningLogger),
                scopeAccessInformation,
                returnRequirements
            );
        }
    }
}
