using CSharpWriter.CodeTranslation.StatementTranslation;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
	public static class ITranslateIndividualStatements_Extensions
	{
		/// <summary>
		/// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
		/// </summary>
		public static TranslatedStatementContentDetails Translate(
			this ITranslateIndividualStatements statementTranslator,
			Statement statement,
			ScopeAccessInformation scopeAccessInformation)
		{
			if (statementTranslator == null)
				throw new ArgumentNullException("statementTranslator");
			if (statement == null)
				throw new ArgumentNullException("statement");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			return Translate(statementTranslator, statement, scopeAccessInformation, ExpressionReturnTypeOptions.None);
		}

		/// <summary>
		/// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
		/// </summary>
		public static TranslatedStatementContentDetails Translate(
			this ITranslateIndividualStatements statementTranslator,
			Expression expression,
			ScopeAccessInformation scopeAccessInformation,
			ExpressionReturnTypeOptions returnRequirements)
		{
			if (statementTranslator == null)
				throw new ArgumentNullException("statementTranslator");
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			return Translate(statementTranslator, (Statement)expression, scopeAccessInformation, returnRequirements);
		}

		private static TranslatedStatementContentDetails Translate(
			ITranslateIndividualStatements statementTranslator,
			Statement statement,
			ScopeAccessInformation scopeAccessInformation,
			ExpressionReturnTypeOptions returnRequirements)
		{
			if (statementTranslator == null)
				throw new ArgumentNullException("statementTranslator");
			if (statement == null)
				throw new ArgumentNullException("statement");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			var expressions = VBScriptTranslator.StageTwoParser.ExpressionParsing.ExpressionGenerator.Generate(statement.BracketStandardisedTokens).ToArray();
			if (expressions.Length != 1)
				throw new ArgumentException("Statement translation should always result in a single expression being generated");

			return statementTranslator.Translate(expressions[0], scopeAccessInformation, returnRequirements);
		}
	}
}
