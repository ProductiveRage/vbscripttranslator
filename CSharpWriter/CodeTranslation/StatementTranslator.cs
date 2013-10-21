using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Misc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using LegacyParser = VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class StatementTranslator : ITranslateIndividualStatements
    {
        private readonly CSharpName _supportClassName;
        private readonly VBScriptNameRewriter _nameRewriter;
        private readonly TempValueNameGenerator _tempNameGenerator;
        public StatementTranslator(CSharpName supportClassName, VBScriptNameRewriter nameRewriter, TempValueNameGenerator tempNameGenerator)
        {
            if (supportClassName == null)
                throw new ArgumentNullException("supportClassName");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (tempNameGenerator == null)
                throw new ArgumentNullException("tempNameGenerator");

            _supportClassName = supportClassName;
            _nameRewriter = nameRewriter;
            _tempNameGenerator = tempNameGenerator;
        }

		/// <summary>
		/// This will never return null or blank, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
		/// </summary>
		public string Translate(LegacyParser.Statement statement)
		{
			if (statement == null)
				throw new ArgumentNullException("statement");

			return Translate(statement, ExpressionReturnTypeOptions.NotSpecified);
		}

		/// <summary>
		/// This will never return null or blank, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
		/// </summary>
		public string Translate(LegacyParser.Expression expression, ExpressionReturnTypeOptions returnRequirements)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			return Translate((LegacyParser.Statement)expression, returnRequirements);
		}

		private string Translate(LegacyParser.Statement statement, ExpressionReturnTypeOptions returnRequirements)
        {
			if (statement == null)
				throw new ArgumentNullException("statement");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			var expressions = ExpressionGenerator.Generate(statement.BracketStandardisedTokens).ToArray();
			if (expressions.Length != 1)
				throw new ArgumentException("Statement translation should always result in a single expression being generated");

			return Translate(expressions[0], ExpressionReturnTypeOptions.NotSpecified);
		}

		private string Translate(Expression expression, ExpressionReturnTypeOptions returnRequirements)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

            // Assert expectations about numbers of segments and operators (if any)
            var segments = expression.Segments.ToArray();
            if (segments.Length == 0)
                throw new ArgumentException("The expression was broken down into zero segments - invalid content");
            if (segments.Length > 3)
                throw new ArgumentException("Expressions with more than three segments are invalid (they must be broken down further), this one has " + segments.Length);
            var operatorSegmentsWithIndexes = segments
                .Select((segment, index) => new { Segment = segment as OperationExpressionSegment, Index = index })
                .Where(s => s.Segment != null);
            if (operatorSegmentsWithIndexes.Count() > 1)
                throw new ArgumentException("Expressions with more than one operators are invalid (they must be broken down further), this one has " + operatorSegmentsWithIndexes.Count());

            // Assert expectations about combinations of segments and operators (if any)
            var operatorSegmentWithIndex = operatorSegmentsWithIndexes.SingleOrDefault();
            if (operatorSegmentWithIndex == null)
            {
                if (segments.Length != 1)
                    throw new ArgumentException("Expressions with multiple segments are not invalid if there is no operator");
            }
            else
            {
                if (segments.Length == 1)
                    throw new ArgumentException("Expressions containing only a single segment, where that segment is an operator, are invalid");
                else if (segments.Length == 2)
                {
                    if (operatorSegmentWithIndex.Index != 0)
                        throw new ArgumentException("If there are any only two segments then the first must be a negation operator (the first here isn't an operator)");
                    if ((operatorSegmentWithIndex.Segment.Token.Content != "-")
                    && !operatorSegmentWithIndex.Segment.Token.Content.Equals("NOT", StringComparison.InvariantCultureIgnoreCase))
                        throw new ArgumentException("If there are any only two segments then the first must be a negation operator (here it has the token content \"" + operatorSegmentWithIndex.Segment.Token.Content + "\")");
                }
                else if (operatorSegmentWithIndex.Index != 1)
                    throw new ArgumentException("If there are three segments, then the middle must be an operator");
            }

			if (segments.Length == 1)
			{
				var result = TranslateNonOperatorSegment(segments[0]);
				return ApplyReturnTypeGuarantee(
					result.Item1,
					result.Item2,
					returnRequirements
				);
			}

            if (segments.Length == 2)
            {
				return ApplyReturnTypeGuarantee(
					string.Format(
						"{0}.{1}({2})",
						_supportClassName.Name,
						GetSupportFunctionName(operatorSegmentWithIndex.Segment.Token),
						TranslateNonOperatorSegment(segments[1])
					),
					ExpressionReturnTypeOptions.Value, // This will be a negation operation and so will always return a numeric value
					returnRequirements
				);
            }

			return ApplyReturnTypeGuarantee(
				string.Format(
					"{0}.{1}({2}, {3})",
					_supportClassName.Name,
					TranslateNonOperatorSegment(segments[0]),
					GetSupportFunctionName(operatorSegmentWithIndex.Segment.Token),
					TranslateNonOperatorSegment(segments[2])
				),
				ExpressionReturnTypeOptions.Value, // All VBScript operators return numeric (or boolean, which are also numeric in VBScript) values
				returnRequirements
            );
        }

        private Tuple<string, ExpressionReturnTypeOptions> TranslateNonOperatorSegment(IExpressionSegment segment)
        {
            if (segment == null)
                throw new ArgumentNullException("segment");
            if (segment is OperationExpressionSegment)
                throw new ArgumentException("This will not accept OperationExpressionSegment instances");

            var numericValueSegment = segment as NumericValueExpressionSegment;
			if (numericValueSegment != null)
				return Tuple.Create(numericValueSegment.Token.Content, ExpressionReturnTypeOptions.Value);

            var stringValueSegment = segment as StringValueExpressionSegment;
            if (stringValueSegment != null)
				return Tuple.Create(stringValueSegment.Token.Content.ToLiteral(), ExpressionReturnTypeOptions.Value);

            var callExpressionSegment = segment as CallExpressionSegment;
            if (callExpressionSegment != null)
                return Translate(callExpressionSegment);

            var callSetExpressionSegment = segment as CallSetExpressionSegment;
            if (callSetExpressionSegment != null)
                return Translate(callSetExpressionSegment);

            var bracketedExpressionSegment = segment as BracketedExpressionSegment;
            if (bracketedExpressionSegment != null)
                return Translate(bracketedExpressionSegment);

            throw new NotImplementedException(); // TODO
        }

		private Tuple<string, ExpressionReturnTypeOptions> Translate(BracketedExpressionSegment bracketedExpressionSegment)
        {
            if (bracketedExpressionSegment == null)
                throw new ArgumentNullException("bracketedExpressionSegment");

            throw new NotImplementedException(); // TODO
        }

        private Tuple<string, ExpressionReturnTypeOptions> Translate(CallExpressionSegment callExpressionSegment)
        {
            if (callExpressionSegment == null)
                throw new ArgumentNullException("callExpressionSegment");

            return TranslateCallExpressionSegment(
                _nameRewriter.GetMemberAccessTokenName(callExpressionSegment.MemberAccessTokens.First()),
                callExpressionSegment.MemberAccessTokens.Skip(1),
                callExpressionSegment.Arguments
            );
        }

        private Tuple<string, ExpressionReturnTypeOptions> TranslateCallExpressionSegment(
            string targetName,
            IEnumerable<IToken> targetMemberAccessTokens,
            IEnumerable<Expression> arguments)
        {
            if (string.IsNullOrWhiteSpace(targetName))
                throw new ArgumentException("Null/blank targetName specified");
            if (targetMemberAccessTokens == null)
                throw new ArgumentNullException("targetMemberAccessTokens");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            // The "master" CALL method signature is
            //
            //   CALL(object target, IEnumerable<string> members, params object[] arguments)
            //
            // (the arguments set is passed as an array as VBScript parameters are, by default, by-ref and so all of the arguments have to be
            // passed in this manner in case any of them need to be access in this manner).
            //
            // However, there are alternate signatures to try to make the most common calls easier to read - eg.
            //
            //   CALL(object)
            //   CALL(object, params object[] arguments)
            //   CALL(object, string member1, params object[] arguments)
            //   CALL(object, string member1, string member2, params object[] arguments)
            //   ..
            //   CALL(object, string member1, string member2, string member3, string member4, string member5, params object[] arguments)
            //
            // The maximum number of member access tokens (one being the initial object and further being the properties / functions accessed)
            // that may use one of these alternate signatures is stored in the constant maxNumberOfMemberAccessorBeforeArraysRequired
            const int maxNumberOfMemberAccessorBeforeArraysRequired = 6;

            // Note: Even if there is only a single member access token (meaning call for "a" not "a.Test") and there are no arguments, we
            // still need to use the CALL method to account for any handling of default properties (eg. a statement "Test" may be a call to
            // a method named "Test" or "Test" may be an instance of a class which has a default parameter-less function, in which case the
            // default function will be executed by that statement.

            var callExpressionContent = new StringBuilder();
            callExpressionContent.AppendFormat(
                "{0}.CALL({1}",
                _supportClassName.Name,
                targetName
            );

            var targetMemberAccessTokensArray = targetMemberAccessTokens.ToArray();
            if (targetMemberAccessTokensArray.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in targetMemberAccessTokens set");

            var ableToUseShorthandCallSignature = (targetMemberAccessTokensArray.Length <= (maxNumberOfMemberAccessorBeforeArraysRequired - 1));
            if (targetMemberAccessTokensArray.Length > 0)
            {
                callExpressionContent.Append(", ");
                if (!ableToUseShorthandCallSignature)
                    callExpressionContent.Append(" new[] { ");
                for (var index = 0; index < targetMemberAccessTokensArray.Length; index++)
                {
                    callExpressionContent.Append(
                        _nameRewriter.GetMemberAccessTokenName(targetMemberAccessTokensArray[index]).ToLiteral()
                    );
                    if (index < (targetMemberAccessTokensArray.Length - 1))
                        callExpressionContent.Append(", ");
                }
                if (!ableToUseShorthandCallSignature)
                    callExpressionContent.Append(" }");
            }

            var argumentsArray = arguments.ToArray();
            if (argumentsArray.Any(a => a == null))
                throw new ArgumentException("Null reference encountered in arguments set");

            if (argumentsArray.Length > 0)
            {
                callExpressionContent.Append(", ");
                if (!ableToUseShorthandCallSignature)
                    callExpressionContent.Append("new object[] { ");
                for (var index = 0; index < argumentsArray.Length; index++)
                {
                    callExpressionContent.Append(
                        Translate(argumentsArray[index], ExpressionReturnTypeOptions.NotSpecified)
                    );
                    if (index < (argumentsArray.Length - 1))
                        callExpressionContent.Append(", ");
                }
                if (!ableToUseShorthandCallSignature)
                    callExpressionContent.Append(" }");
            }

            callExpressionContent.Append(")");
			return Tuple.Create(
				callExpressionContent.ToString(),
				ExpressionReturnTypeOptions.NotSpecified // This could be anything so we have to report NotSpecified as the return type
			);
        }

        private Tuple<string, ExpressionReturnTypeOptions> Translate(CallSetExpressionSegment callSetExpressionSegment)
        {
            if (callSetExpressionSegment == null)
                throw new ArgumentNullException("callSetExpressionSegment");
            
            var content = "";
            var numberOfCallExpressions = callSetExpressionSegment.CallExpressionSegments.Count();
            for (var index = 0; index < numberOfCallExpressions; index++)
            {
                var callExpression = callSetExpressionSegment.CallExpressionSegments.ElementAt(index);
                if (index == 0)
                {
                    content = Translate(callExpression).Item1;
                    continue;
                }
                content = TranslateCallExpressionSegment(
                    content,
                    callExpression.MemberAccessTokens,
                    callExpression.Arguments
                ).Item1;
            }
            return Tuple.Create(
                content,
                ExpressionReturnTypeOptions.NotSpecified // This could be anything so we have to report NotSpecified as the return type
            );
        }

        private string GetSupportFunctionName(OperatorToken operatorToken)
        {
            if (operatorToken == null)
                throw new ArgumentNullException("operatorToken");

            switch (operatorToken.Content.ToUpper())
            {
                // Arithmetic operators
                case "^":   return "POW";
                case "/":   return "DIV";
                case "*":   return "MULT";
                case "\\":  return "INTDIV";
                case "MOD": return "MOD";
                case "+":   return "ADD";
                case "-":   return "SUBT";

                // String concatenation
                case "&":   return "CONCAT";

                // Logical operators
                case "NOT": return "NOT";
                case "AND": return "AND";
                case "OR":  return "OR";
                case "XOR": return "XOR";

                // Comparison operators
                case "=":   return "EQ";
                case "<>":  return "NOTEQ";
                case "<":   return "LT";
                case ">":   return "GT";
                case "<=":  return "LTE";
                case ">=":  return "GTE";
                case "IS":  return "IS";
                case "EQV": return "EQV";
                case "IMP": return "IMP";

                // Error
                default:
                    throw new NotSupportedException("Unsupported OperatorToken content: " + operatorToken.Content);
            }
        }

		private string ApplyReturnTypeGuarantee(string translatedContent, ExpressionReturnTypeOptions contentType, ExpressionReturnTypeOptions requiredReturnType)
		{
			if (string.IsNullOrWhiteSpace(translatedContent))
				throw new ArgumentException("Null/blank translatedContent specified");

			switch(requiredReturnType)
			{
				case ExpressionReturnTypeOptions.Boolean:
					return string.Format(
						"{0}.IF({1})",
						_supportClassName.Name,
						translatedContent
					);

				case ExpressionReturnTypeOptions.NotSpecified:
					return translatedContent;

				case ExpressionReturnTypeOptions.Reference:
					// If we know that this returns a value type then we can tell at this point that it's not going to work. If it returns a Reference
					// type then we're golden. If contentType is NotSpecified then we just have to hope for the best that it's appropriate.
					if ((contentType == ExpressionReturnTypeOptions.Boolean) || (contentType == ExpressionReturnTypeOptions.Value))
						throw new ArgumentException("Invalid content, Expression specified needs to return an Reference (Object), but returns " + contentType + " (this would result in a compile-time \"Object Expected\")");
					return translatedContent;

				case ExpressionReturnTypeOptions.Value:
					return string.Format(
						"{0}.VAL(1})",
						_supportClassName.Name,
						translatedContent
					);

				default:
					throw new NotSupportedException("Unsupported requiredReturnType value: " + requiredReturnType);
			}
		}
    }
}
