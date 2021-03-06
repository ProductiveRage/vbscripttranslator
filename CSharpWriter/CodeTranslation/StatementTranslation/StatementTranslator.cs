﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions;
using VBScriptTranslator.CSharpWriter.Lists;
using VBScriptTranslator.CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation
{
	public class StatementTranslator : ITranslateIndividualStatements
	{
		private readonly CSharpName _supportRefName, _envRefName, _outerRefName;
		private readonly VBScriptNameRewriter _nameRewriter;
		private readonly TempValueNameGenerator _tempNameGenerator;
		private readonly ILogInformation _logger;
		public StatementTranslator(
			CSharpName supportRefName,
			CSharpName envRefName,
			CSharpName outerRefName,
			VBScriptNameRewriter nameRewriter,
			TempValueNameGenerator tempNameGenerator,
			ILogInformation logger)
		{
			if (supportRefName == null)
				throw new ArgumentNullException("supportRefName");
			if (nameRewriter == null)
				throw new ArgumentNullException("nameRewriter");
			if (outerRefName == null)
				throw new ArgumentNullException("outerRefName");
			if (tempNameGenerator == null)
				throw new ArgumentNullException("tempNameGenerator");
			if (logger == null)
				throw new ArgumentNullException("logger");

			_supportRefName = supportRefName;
			_envRefName = envRefName;
			_outerRefName = outerRefName;
			_nameRewriter = nameRewriter;
			_tempNameGenerator = tempNameGenerator;
			_logger = logger;
		}

		/// <summary>
		/// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null expression reference)
		/// </summary>
		public TranslatedStatementContentDetails Translate(Expression expression, ScopeAccessInformation scopeAccessInformation, ExpressionReturnTypeOptions returnRequirements)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			// See notes in TryToGetShortCutStatementResponse method..
			var shortCutStatementResponse = TryToGetShortCutStatementResponse(expression, scopeAccessInformation, returnRequirements);
			if (shortCutStatementResponse != null)
				return shortCutStatementResponse;

			// See notes in TryToGetConcatFlattenedSpecialCaseResponse method..
			var concatFlattenedSpecialCaseResponse = TryToGetConcatFlattenedSpecialCaseResponse(expression, scopeAccessInformation, returnRequirements);
			if (concatFlattenedSpecialCaseResponse != null)
				return concatFlattenedSpecialCaseResponse;

			// Assert expectations about numbers of segments and operators (if any)
			// - There may not be more than three segments, and only three where there are two values or calls separated by an operator. CallSetExpressionSegments and
			//   BracketedExpressionSegments are key to ensuring that this format is met.
			var segments = expression.Segments.ToArray();
			if (segments.Length == 0)
				throw new ArgumentException("The expression was broken down into zero segments - invalid content");
			if (segments.Length > 3)
				throw new ArgumentException("Expressions with more than three segments are invalid (they must be processed further, potentially using CallSetExpressionSegments and BracketedExpressionSegments where appropriate), this one has " + segments.Length);
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
					throw new ArgumentException("Expressions with multiple segments are not invalid if there is no operator (line " + (segments.First().AllTokens.First().LineIndex + 1) + ")");
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
				var result = TranslateNonOperatorSegment(segments[0], scopeAccessInformation);
				return new TranslatedStatementContentDetails(
					ApplyReturnTypeGuarantee(
						result.TranslatedContent,
						result.ContentType,
						returnRequirements,
						segments[0].AllTokens.First().LineIndex
					),
					result.VariablesAccessed
				);
			}

			if (segments.Length == 2)
			{
				var result = TranslateNonOperatorSegment(segments[1], scopeAccessInformation);
				return new TranslatedStatementContentDetails(
					ApplyReturnTypeGuarantee(
						string.Format(
							"{0}.{1}({2})",
							_supportRefName.Name,
							GetSupportFunctionName(operatorSegmentWithIndex.Segment.Token),
							result.TranslatedContent
						),
						ExpressionReturnTypeOptions.Value, // This will be a negation operation and so will always return a numeric value
						returnRequirements,
						segments[0].AllTokens.First().LineIndex
					),
					result.VariablesAccessed
				);
			}

			var segmentLeft = segments[0];
			var segmentRight = segments[2];
			bool mustConvertLeftValueToNumber, mustConvertRightValueToNumber;
			bool mustConvertLeftValueToDate, mustConvertRightValueToDate;
			bool mustConvertLeftValueToString, mustConvertRightValueToString;
			if (operatorSegmentWithIndex.Segment.Token is ComparisonOperatorToken)
			{
				// If the operator segment is a ComparisonOperatorToken (not a LogicalOperatorToken and not any other type of OperatorToken, such
				// as an arithmetic operation or string concatenation) then there are special rules to apply if one side of the operation is known
				// to be a constant at compile time - namely, the other side must be parseable as a numeric value otherwise a "Type mismatch" will
				// be raised. If both sides of the operation are variables and one of them is a numeric value and the other a word, then the
				// comparison will return false but it won't raise an error. There is similar logic for string constants (but no equivalent
				// for boolean constants).
				// See http://blogs.msdn.com/b/ericlippert/archive/2004/07/30/202432.aspx for details about "hard types" that pertain to this
				// - Update: This only applies to numeric values if they are non-negative. It also doesn't apply if they WOULD be non-negative
				//   if multiple negative signs were remove - eg. the following conditions do NOT count as containing numeric literals:
				//     If ("a" = -1) Then
				//     If ("a" = --1) Then
				//     If ("a" = +1) Then
				//   The first is a negative number, the second looks like it could be considered to be a positive number but the VBScript
				//   interpreter will not cancel out those double negative signs and still consider it a literal. When source content is
				//   being parsed, this must be considered (currently the OperatorCombiner will replace --1 with CDbl(1) so that it's
				//   obvious to the processing here that it should not be a numeric literal - if it replaced --1 with 1 then it WOULD
				//   look like a numeric literal here and there would be an inconsistency with the VBScript interpreter).
				// - Further update: Eric didn't talk about date literals, but they also have similar behaviour to numbers and strings! :(
				//   Dates have higher priority than strings but lower than numbers, so the following is an error
				//     If ("a" = #2015-5-27#) Then
				//   because "a" can not be parsed into a date, but
				//     If (-657435 = #2015-5-27#) Then
				//   is not an error, even though -657435 falls just outside of the VBScript date ranges (because #2015-5-27# is interpreted as
				//   a number, rather than -657435 being parsed as a date)
				var segmentLeftAsNumericValue = segmentLeft as NumericValueExpressionSegment;
				var segmentRightAsNumericValue = segmentRight as NumericValueExpressionSegment;
				var segmentLeftIsNonNegativeNumericValue = (segmentLeftAsNumericValue != null) && (segmentLeftAsNumericValue.Token.Value >= 0);
				var segmentRightIsNonNegativeNumericValue = (segmentRightAsNumericValue != null) && (segmentRightAsNumericValue.Token.Value >= 0);
				if ((segmentLeftAsNumericValue != null) && (segmentRightAsNumericValue != null))
				{
					// If both sides of an operation are numeric constants, then the comparison will be simple and not require any interfering in
					// terms of casting values at this point
					mustConvertLeftValueToNumber = false;
					mustConvertRightValueToNumber = false;
					mustConvertLeftValueToDate = false;
					mustConvertRightValueToDate = false;
					mustConvertLeftValueToString = false;
					mustConvertRightValueToString = false;
				}
				else if (segmentLeftIsNonNegativeNumericValue || segmentRightIsNonNegativeNumericValue)
				{
					// However, if one side of an operation is a numeric constant, then the other side must be parseable as a number - otherwise
					// a "Type mismatch" error should be raised; the statement "IF ("aa" > 0) THEN" will error, for example
					mustConvertLeftValueToNumber = !segmentLeftIsNonNegativeNumericValue;
					mustConvertRightValueToNumber = !segmentRightIsNonNegativeNumericValue;
					mustConvertLeftValueToDate = false;
					mustConvertRightValueToDate = false;
					mustConvertLeftValueToString = false;
					mustConvertRightValueToString = false;
				}
				else if ((segmentLeft is DateValueExpressionSegment) && (segmentRight is DateValueExpressionSegment))
				{
					// If both sides of an operation are date literals, then the comparison will be simple and not require any interfering in
					// terms of casting values at this point
					mustConvertLeftValueToNumber = false;
					mustConvertRightValueToNumber = false;
					mustConvertLeftValueToDate = false;
					mustConvertRightValueToDate = false;
					mustConvertLeftValueToString = false;
					mustConvertRightValueToString = false;
				}
				else if ((segmentLeft is DateValueExpressionSegment) || (segmentRight is DateValueExpressionSegment))
				{
					// However, if one side of an operation is a date literal then the other side must be parsed as a date before the comparison
					// is made
					mustConvertLeftValueToNumber = false;
					mustConvertRightValueToNumber = false;
					mustConvertLeftValueToDate = !(segmentLeft is DateValueExpressionSegment);
					mustConvertRightValueToDate = !(segmentRight is DateValueExpressionSegment);
					mustConvertLeftValueToString = false;
					mustConvertRightValueToString = false;
				}
				else if ((segmentLeft is StringValueExpressionSegment) && (segmentRight is StringValueExpressionSegment))
				{
					// If both sides of an operation are string constants, then the comparison will be simple and not require any interfering in
					// terms of casting values at this point
					mustConvertLeftValueToNumber = false;
					mustConvertRightValueToNumber = false;
					mustConvertLeftValueToDate = false;
					mustConvertRightValueToDate = false;
					mustConvertLeftValueToString = false;
					mustConvertRightValueToString = false;
				}
				else if ((segmentLeft is StringValueExpressionSegment) || (segmentRight is StringValueExpressionSegment))
				{
					// However, if one side of an operation is a string constant then the other side will be parsed as a string before the
					// comparison is made. This actually makes it more permissive, rather than more constricting (which the numeric value
					// coercing does). For example,
					//   If ("True" = true) Then
					// will match as the boolean true will be converted into the string "True" and this will match the constant. Unlike
					//   Dim vStringTrue: vStringTrue = "True"
					//   If (vStringTrue = true) Then
					// which will NOT match as the string value "True" does not match the boolean value true and the boolean value is not
					// converted into a string since there is no string constant in the statement.
					mustConvertLeftValueToNumber = false;
					mustConvertRightValueToNumber = false;
					mustConvertLeftValueToDate = false;
					mustConvertRightValueToDate = false;
					mustConvertLeftValueToString = !(segmentLeft is StringValueExpressionSegment);
					mustConvertRightValueToString = !(segmentRight is StringValueExpressionSegment);
				}
				else
				{
					// If neither side is a literal then there is nothing to worry about (any complexities of how values should be compared
					// can be left up to the runtime comparison method implementation)
					mustConvertLeftValueToNumber = false;
					mustConvertRightValueToNumber = false;
					mustConvertLeftValueToDate = false;
					mustConvertRightValueToDate = false;
					mustConvertLeftValueToString = false;
					mustConvertRightValueToString = false;
				}
			}
			else
			{
				mustConvertLeftValueToNumber = false;
				mustConvertRightValueToNumber = false;
				mustConvertLeftValueToDate = false;
				mustConvertRightValueToDate = false;
				mustConvertLeftValueToString = false;
				mustConvertRightValueToString = false;
			}
			var resultLeft = TranslateNonOperatorSegment(segmentLeft, scopeAccessInformation);
			if (new[] { mustConvertLeftValueToNumber, mustConvertLeftValueToDate, mustConvertLeftValueToString}.Count(v => v) > 1)
				throw new Exception("Something went wrong in the processing, no more than one of mustConvertLeftValueToNumber, mustConvertLeftValueToDate and mustConvertLeftValueToString may be set for a comparison");
			if (mustConvertLeftValueToNumber)
				resultLeft = WrapTranslatedResultInNumericCast(resultLeft);
			else if (mustConvertLeftValueToDate)
				resultLeft = WrapTranslatedResultInDateCast(resultLeft);
			else if (mustConvertLeftValueToString)
				resultLeft = WrapTranslatedResultInStringConversion(resultLeft);
			var resultRight = TranslateNonOperatorSegment(segmentRight, scopeAccessInformation);
			if (new[] { mustConvertRightValueToNumber, mustConvertRightValueToDate, mustConvertRightValueToString }.Count(v => v) > 1)
				throw new Exception("Something went wrong in the processing, no more than one of mustConvertRightValueToNumber, mustConvertRightValueToDate and mustConvertRightValueToString may be set for a comparison");
			if (mustConvertRightValueToNumber)
				resultRight = WrapTranslatedResultInNumericCast(resultRight);
			else if (mustConvertRightValueToDate)
				resultRight = WrapTranslatedResultInDateCast(resultRight);
			else if (mustConvertRightValueToString)
				resultRight = WrapTranslatedResultInStringConversion(resultRight);
			return new TranslatedStatementContentDetails(
				ApplyReturnTypeGuarantee(
					string.Format(
						"{0}.{1}({2}, {3})",
						_supportRefName.Name,
						GetSupportFunctionName(operatorSegmentWithIndex.Segment.Token),
						resultLeft.TranslatedContent,
						resultRight.TranslatedContent
					),
					ExpressionReturnTypeOptions.Value, // All VBScript operators return numeric (or boolean, which are also numeric in VBScript) values
					returnRequirements,
					segments[0].AllTokens.First().LineIndex
				),
				resultLeft.VariablesAccessed.AddRange(resultRight.VariablesAccessed)
			);
		}

		private TranslatedStatementContentDetailsWithContentType WrapTranslatedResultInNumericCast(TranslatedStatementContentDetailsWithContentType translatedStatement)
		{
			if (translatedStatement == null)
				throw new ArgumentNullException("translatedStatement");

			// Note: We have to use NullableNUM here rather than NUM since VBScript Null is acceptable in cases such as
			//   If (Null = 1) Then
			// but it is not in other cases where NUM must be used such as
			//   For i = Null To 10
			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"{0}.NullableNUM({1})",
					_supportRefName.Name,
					translatedStatement.TranslatedContent
				),
				ExpressionReturnTypeOptions.Value,
				translatedStatement.VariablesAccessed
			);
		}

		private TranslatedStatementContentDetailsWithContentType WrapTranslatedResultInDateCast(TranslatedStatementContentDetailsWithContentType translatedStatement)
		{
			if (translatedStatement == null)
				throw new ArgumentNullException("translatedStatement");

			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"{0}.NullableDATE({1})",
					_supportRefName.Name,
					translatedStatement.TranslatedContent
				),
				ExpressionReturnTypeOptions.Value,
				translatedStatement.VariablesAccessed
			);
		}

		private TranslatedStatementContentDetailsWithContentType WrapTranslatedResultInStringConversion(TranslatedStatementContentDetailsWithContentType translatedStatement)
		{
			if (translatedStatement == null)
				throw new ArgumentNullException("translatedStatement");

			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"{0}.NullableSTR({1})",
					_supportRefName.Name,
					translatedStatement.TranslatedContent
				),
				ExpressionReturnTypeOptions.Value,
				translatedStatement.VariablesAccessed
			);
		}

		private TranslatedStatementContentDetailsWithContentType TranslateNonOperatorSegment(IExpressionSegment segment, ScopeAccessInformation scopeAccessInformation)
		{
			if (segment == null)
				throw new ArgumentNullException("segment");
			if (segment is OperationExpressionSegment)
				throw new ArgumentException("This will not accept OperationExpressionSegment instances");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var numericValueSegment = segment as NumericValueExpressionSegment;
			if (numericValueSegment != null)
			{
				return new TranslatedStatementContentDetailsWithContentType(
					numericValueSegment.Token.AsCSharpValue(),
					ExpressionReturnTypeOptions.Value,
					new NonNullImmutableList<NameToken>()
				);
			}

			var stringValueSegment = segment as StringValueExpressionSegment;
			if (stringValueSegment != null)
			{
				return new TranslatedStatementContentDetailsWithContentType(
					stringValueSegment.Token.Content.ToLiteral(),
					ExpressionReturnTypeOptions.Value,
					new NonNullImmutableList<NameToken>()
				);
			}

			var dateValueSegment = segment as DateValueExpressionSegment;
			if (dateValueSegment != null)
			{
				// When initialising a date literal, the culture at runtime may affect the value, as may the current year if the format of the literal did not specify
				// a year. As such, date literals are interpreted at runtime. In the case of a "dynamic year" date literal, the compat layer will ensure that the year
				// from when the request started is used in all cases - if the request is slow and the year ticks over part-way through, all date literals with dynamic
				// years must be use the year from when the request started. Subsequent requests will use the new year. This is consistent with VBScript since the
				// script is re-interpreted each time it's run.
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"{0}.DateLiteralParser.Parse({1})",
						_supportRefName.Name,
						dateValueSegment.Token.Content.ToLiteral()
					),
					ExpressionReturnTypeOptions.Value,
					new NonNullImmutableList<NameToken>()
				);
			}

			var builtInValueExpressionSegment = segment as BuiltInValueExpressionSegment;
			if (builtInValueExpressionSegment != null)
				return Translate(builtInValueExpressionSegment);

			var callExpressionSegment = segment as CallExpressionSegment;
			if (callExpressionSegment != null)
				return Translate(callExpressionSegment, scopeAccessInformation);

			var callSetExpressionSegment = segment as CallSetExpressionSegment;
			if (callSetExpressionSegment != null)
				return Translate(callSetExpressionSegment, scopeAccessInformation);

			var bracketedExpressionSegment = segment as BracketedExpressionSegment;
			if (bracketedExpressionSegment != null)
				return Translate(bracketedExpressionSegment, scopeAccessInformation);

			var newInstanceExpressionSegment = segment as NewInstanceExpressionSegment;
			if (newInstanceExpressionSegment != null)
				return Translate(newInstanceExpressionSegment, scopeAccessInformation.ScopeDefiningParent.Scope);

			var runtimeErrorExpressionSegment = segment as RuntimeErrorExpressionSegment;
			if (runtimeErrorExpressionSegment != null)
				return Translate(runtimeErrorExpressionSegment);

			throw new NotSupportedException("Unsupported segment type: " + segment.GetType());
		}

		private TranslatedStatementContentDetailsWithContentType Translate(BracketedExpressionSegment bracketedExpressionSegment, ScopeAccessInformation scopeAccessInformation)
		{
			if (bracketedExpressionSegment == null)
				throw new ArgumentNullException("bracketedExpressionSegment");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			// 2014-12-08 DWR: This previously wrapped the returned content in brackets - largely only because the source is a bracketed expression. But
			// since they're always broken down to respect VBScript's operator precedence and then passed through functions for every operation, there
			// is no benefit to adding further bracketing, so it's been removed.
			var translatedInnerContentDetails = Translate(
				new Expression(bracketedExpressionSegment.Segments),
				scopeAccessInformation,
				ExpressionReturnTypeOptions.NotSpecified
			);
			return new TranslatedStatementContentDetailsWithContentType(
				translatedInnerContentDetails.TranslatedContent,
				ExpressionReturnTypeOptions.NotSpecified,
				translatedInnerContentDetails.VariablesAccessed
			);
		}

		private TranslatedStatementContentDetailsWithContentType Translate(BuiltInValueExpressionSegment builtInValueExpressionSegment)
		{
			if (builtInValueExpressionSegment == null)
				throw new ArgumentNullException("builtInValueExpressionSegment");

			// Handle non-constants special cases
			if (builtInValueExpressionSegment.Token.Content.Equals("err", StringComparison.InvariantCultureIgnoreCase))
			{
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"{0}.ERR",
						_supportRefName.Name
					),
					ExpressionReturnTypeOptions.Reference,
					new NonNullImmutableList<NameToken>()
				);
			}

			// Handle constants special cases
			if (builtInValueExpressionSegment.Token.Content.Equals("nothing", StringComparison.OrdinalIgnoreCase))
			{
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"VBScriptConstants.Nothing",
						_supportRefName.Name
					),
					ExpressionReturnTypeOptions.Reference,
					new NonNullImmutableList<NameToken>()
				);
			}
			else if (builtInValueExpressionSegment.Token.Content.Equals("true", StringComparison.OrdinalIgnoreCase))
			{
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"true",
						_supportRefName.Name
					),
					ExpressionReturnTypeOptions.Boolean,
					new NonNullImmutableList<NameToken>()
				);
			}
			else if (builtInValueExpressionSegment.Token.Content.Equals("false", StringComparison.OrdinalIgnoreCase))
			{
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"false",
						_supportRefName.Name
					),
					ExpressionReturnTypeOptions.Boolean,
					new NonNullImmutableList<NameToken>()
				);
			}

			// Handle regular value-type constants
			var constantProperty = typeof(VBScriptConstants).GetProperty(
				builtInValueExpressionSegment.Token.Content,
				BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Static
			);
			if ((constantProperty == null) || !constantProperty.CanRead || constantProperty.GetIndexParameters().Any())
				throw new NotSupportedException("Unsupported BuiltInValueToken content: " + builtInValueExpressionSegment.Token.Content);
			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"VBScriptConstants.{1}",
					_supportRefName.Name,
					constantProperty.Name
				),
				ExpressionReturnTypeOptions.Value,
				new NonNullImmutableList<NameToken>()
			);
		}

		/// <summary>
		/// This may only be called when a CallExpressionSegment is encountered as one of the segments in the Expression passed into the public Translate method
		/// or if it is the first segment in a CallSetExpressionSegment, subsequent segments in a CallSetExpressionSegment should be passed direct into the 
		/// TranslateCallExpressionSegment method)
		/// </summary>
		private TranslatedStatementContentDetailsWithContentType Translate(CallExpressionSegment callExpressionSegment, ScopeAccessInformation scopeAccessInformation)
		{
			if (callExpressionSegment == null)
				throw new ArgumentNullException("callExpressionSegment");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			// We may have to monkey about with the data here - if there are references to the return value of the current function (if we're in one) then these
			// need to be replaced with the scopeAccessInformation's parentReturnValueNameIfAny value (if there is one).
			var firstMemberAccessToken = callExpressionSegment.MemberAccessTokens.First();
			var firstMemberAccessTokenIsParentReturnValueName = false;
			if (scopeAccessInformation.ParentReturnValueNameIfAny != null)
			{
				// If the segment's first (or only) member accessor has no arguments and wasn't expressed in the source code as a function call (ie. it didn't
				// have brackets after the member accessor) and the single member accessor matches the name of the ScopeDefiningParent then we need to make
				// the replacement..
				// - If arguments are specified or brackets used with no arguments then it is always a function (or property) call and the return-value
				//   replacement does not need to be made (the return value may not be DIM'd or REDIM'd to an array and so element access is not allowed,
				//   so ANY argument use always points to a function / property call)
				var rewrittenFirstMemberAccessor = _nameRewriter.GetMemberAccessTokenName(firstMemberAccessToken);
				var rewrittenScopeDefiningParentName = _nameRewriter.GetMemberAccessTokenName(scopeAccessInformation.ScopeDefiningParent.Name);
				if ((rewrittenFirstMemberAccessor == rewrittenScopeDefiningParentName)
				&& !callExpressionSegment.Arguments.Any()
				&& (callExpressionSegment.ZeroArgumentBracketsPresence == CallExpressionSegment.ArgumentBracketPresenceOptions.Absent))
				{
					// The ScopeDefiningParent's Name will have come from a TempValueNameGenerator rather than VBScript source code, and as such it should
					// not be passed through any VBScriptNameRewriter processing. Using a DoNotRenameNameToken means that, if the extension method
					// GetMemberAccessTokenName is consistently used for VBScriptNameRewriter access, its name won't be altered.
					var parentReturnValueNameToken = new DoNotRenameNameToken(
						scopeAccessInformation.ParentReturnValueNameIfAny.Name,
						firstMemberAccessToken.LineIndex
					);
					callExpressionSegment = new CallExpressionSegment(
						new[] { parentReturnValueNameToken }.Concat(callExpressionSegment.MemberAccessTokens.Skip(1)),
						new Expression[0],
						CallExpressionSegment.ArgumentBracketPresenceOptions.Absent
					);
					firstMemberAccessToken = parentReturnValueNameToken;
					firstMemberAccessTokenIsParentReturnValueName = true;
				}
			}

			var targetBuiltInFunction = firstMemberAccessToken as BuiltInFunctionToken;
			if (targetBuiltInFunction != null)
			{
				var supportFunctionDetails = GetDetailsOfBuiltInFunction(targetBuiltInFunction, callExpressionSegment.Arguments.Count());
				var rewrittenMemberAccessTokens = new[] { new DoNotRenameNameToken(supportFunctionDetails.SupportFunctionName, targetBuiltInFunction.LineIndex) }
					.Concat(callExpressionSegment.MemberAccessTokens.Skip(1));

				// If the call expression is a single-argument call to "CDbl" and the argument is a numeric literal that VBScript would declare as "Double",
				// we can skip the call entirely. Likewise for "CInt" / "Integer" and "CLng" / "Long". The OperatorCombiner identifies some cases where what
				// appear to be numeric literals must not be treated as numeric literal later on in the process, so it wraps them in a "CInt" / "CLng" / "CDbl"
				// call (whichever is appropriate) - this code allows us to remove those injected functions right at the very last minute without changing the
				// meaning of the translate code. (Note: It is elsewhere in this class where it is important whether tokens should be given "special treatment"
				// as number literals in comparisons or not).
				if ((callExpressionSegment.Arguments.Count() == 1) && (callExpressionSegment.Arguments.Single().Segments.Count() == 1))
				{
					var singleArgumentSegment = callExpressionSegment.Arguments.Single().Segments.Single();
					var singleArgumentSegmentAsNumericValue = singleArgumentSegment as NumericValueExpressionSegment;
					if ((singleArgumentSegmentAsNumericValue != null) && singleArgumentSegmentAsNumericValue.Token.IsSafeToUnwrapFrom(targetBuiltInFunction))
						return TranslateNonOperatorSegment(singleArgumentSegmentAsNumericValue, scopeAccessInformation);
				}

				// If supportFunctionDetails.DesiredNumberOfArgumentsMatchedAgainst is not null then there is a support function that has the same number of
				// arguments as this callExpressionSegment, all of which as not output parameters and not ref parameters and all of which are of type "object".
				// This means that we can express this more directly, without relying upon an "argument provider" - eg.
				//   _.CDate(value)
				// instead of
				//   _.CALL(this, _, "CDate", _.ARGS.VAL(value))
				// If these conditions are not met then we have to use the argument provider - if, for example, the support function requires a string parameter
				// then we can't be sure that the argument we have at this point is a string. If there is no support function that matches the segment's number
				// of arguments then this must fail at runtime, not compile time, and so the "CALL" method approach is required.
				if (supportFunctionDetails.DesiredNumberOfArgumentsMatchedAgainst != null)
					return TranslateAsDirectSupportFunctionCall(supportFunctionDetails, callExpressionSegment.Arguments, scopeAccessInformation);

				return TranslateCallExpressionSegment(
					new DoNotRenameNameToken(_supportRefName.Name, firstMemberAccessToken.LineIndex),
					rewrittenMemberAccessTokens,
					callExpressionSegment.Arguments,
					callExpressionSegment.ZeroArgumentBracketsPresence,
					scopeAccessInformation,
					indexInCallSet: 0, // Since this is a single CallExpressionSegment the indexInCallSet value to pass is always zero
					targetIsKnownToBeBuiltInFunction: true
				);
			}

			// The only BuiltInValueToken that is acceptable here is "Err", since it is the only one that has members that may be accessed. If any other builtin
			// value is accessed in this manner then an "Object required" error needs to be raised at runtime (this will be handled by the code that is generated
			// since it will likely have use the support library's CALL function which will realise that there is no way to access a property on a value type)
			var targetBuiltInValue = firstMemberAccessToken as BuiltInValueToken;
			var targetIsErrReference = (targetBuiltInValue != null) && targetBuiltInValue.Content.Equals("Err", StringComparison.OrdinalIgnoreCase);

			// If the target reference IS "Err", then there is a special case of "Err.Raise" to consider. If Err.Raise is called with the correct number of arguments
			// then it may be mapped directly onto the support function (either 1, 2 or 3 arguments must be present). If a different number of arguments are present
			// then the target still needs rewriting from "Err.Raise" to "_.RAISEERROR", but it will have to go through the CALL function. Same goes for "Err.Clear".
			TranslatedStatementContentDetailsWithContentType specialErrorHandlingFunctionStatementIfApplicable;
			NameToken target;
			var memberAccessors = callExpressionSegment.MemberAccessTokens.Skip(1);
			if (targetIsErrReference)
			{
				string specialErrorHandlingFunctionNameIfApplicable;
				if ((memberAccessors.Count() == 1) && (memberAccessors.Single() is NameToken))
				{
					if (memberAccessors.Single().Content.Equals("RAISE", StringComparison.OrdinalIgnoreCase))
						specialErrorHandlingFunctionNameIfApplicable = "RAISEERROR";
					else if (memberAccessors.Single().Content.Equals("CLEAR", StringComparison.OrdinalIgnoreCase))
						specialErrorHandlingFunctionNameIfApplicable = "CLEARANYERROR";
					else
						specialErrorHandlingFunctionNameIfApplicable = null;
				}
				else
					specialErrorHandlingFunctionNameIfApplicable = null;
				if (specialErrorHandlingFunctionNameIfApplicable != null)
				{
					var specialErrorHandlingFunctionToken = new NameToken(specialErrorHandlingFunctionNameIfApplicable, memberAccessors.Single().LineIndex);
					var errorSupportFunction = GetDetailsOfBuiltInFunction(specialErrorHandlingFunctionToken, callExpressionSegment.Arguments.Count());
					if (errorSupportFunction.DesiredNumberOfArgumentsMatchedAgainst != null)
						specialErrorHandlingFunctionStatementIfApplicable = TranslateAsDirectSupportFunctionCall(errorSupportFunction, callExpressionSegment.Arguments, scopeAccessInformation);
					else
					{
						// Can't call the function directly (argument mismatch that will cause a compile error - so we need to use CALL, which will push the error
						// down to runtime) but we still need to rewrite the request from "Err.Raise" to "_.RAISEERROR" or "Err.Clear" to "_.CLEARERROR"
						specialErrorHandlingFunctionStatementIfApplicable = null;
						memberAccessors = new[] { specialErrorHandlingFunctionToken };
					}
					target = new DoNotRenameNameToken(_supportRefName.Name, firstMemberAccessToken.LineIndex);
				}
				else
				{
					target = new DoNotRenameNameToken(_supportRefName.Name + ".ERR", firstMemberAccessToken.LineIndex);
					specialErrorHandlingFunctionStatementIfApplicable = null;
				}
			}
			else
			{
				// Since the CallExpressionSegment's MemberAccessTokens property is documented as only containing BuiltInFunctionToken, BuiltInValueToken, KeyWordToken and
				// NameToken values, it's safe to assume that it is a NameToken in this case since the other possibilities should have been accounted for by this point.
				target = (NameToken)firstMemberAccessToken;
				specialErrorHandlingFunctionStatementIfApplicable = null;
			}

			// If the target is a function then we can't have that as the target reference in the generated CALL statement, the owner of that function must be the target and
			// the function name must become one of the member accessors (eg. GetSomething().Name can not be represented by _.CALL(GetSomething, "Name") since that is not
			// valid C#, however it CAN be represented by _.CALL(this, _outer, "GetSomething", "Name") or _.CALL(this, "GetSomething", "Name"), depending upon where the function
			// is defined).
			var targetReferenceDetails = scopeAccessInformation.TryToGetDeclaredReferenceDetails(target, _nameRewriter);
			if (targetReferenceDetails != null)
			{
				if (targetReferenceDetails.ReferenceType == ReferenceTypeOptions.Class)
					throw new ArgumentException("Invalid CallExpressionSegment, target is a Class (\"" + target.Content + "\") - not an instance of a class but the class itself");
				if ((targetReferenceDetails.ReferenceType == ReferenceTypeOptions.Function) || (targetReferenceDetails.ReferenceType == ReferenceTypeOptions.Property))
				{
					memberAccessors = new[] { target }.Concat(memberAccessors);
					if (targetReferenceDetails.ScopeLocation == VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions.OutermostScope)
						target = new ProcessedNameToken(_outerRefName.Name, target.LineIndex);
					else
						target = new ProcessedNameToken("this", target.LineIndex);
				}
			}

			TranslatedStatementContentDetailsWithContentType result;
			if (specialErrorHandlingFunctionStatementIfApplicable != null)
				result = specialErrorHandlingFunctionStatementIfApplicable;
			else
			{
				result = TranslateCallExpressionSegment(
					target,
					memberAccessors,
					callExpressionSegment.Arguments,
					callExpressionSegment.ZeroArgumentBracketsPresence,
					scopeAccessInformation,
					indexInCallSet: 0, // Since this is a single CallExpressionSegment the indexInCallSet value to pass is always zero
					targetIsKnownToBeBuiltInFunction: targetIsErrReference // Don't try to rewrite the target reference if it's the Err reference, we've already got it correct
				);
			}
			var targetNameToken = firstMemberAccessToken as NameToken;
			if ((targetNameToken != null) && !firstMemberAccessTokenIsParentReturnValueName)
			{
				result = new TranslatedStatementContentDetailsWithContentType(
					result.TranslatedContent,
					result.ContentType,
					result.VariablesAccessed.Add(targetNameToken)
				);
			}
			return result;
		}

		/// <summary>
		/// This will try to return information about a built in function that matches the name of the specified builtInFunctionToken. It will also try to match
		/// a method signature that has parameters that are all of type object and that are not out or ref - if it identifies such a signature where the number
		/// of these paramaters matches the specified desiredNumberOfArguments, then the return value will have a DesiredNumberOfArgumentsMatchedAgainst value
		/// that is consistent with desiredNumberOfArguments, otherwise it will be null. This information affects how the support function may be called.
		/// </summary>
		private BuiltInFunctionDetails GetDetailsOfBuiltInFunction(IToken builtInFunctionToken, int desiredNumberOfArguments)
		{
			if (builtInFunctionToken == null)
				throw new ArgumentNullException("builtInFunctionToken");
			if (desiredNumberOfArguments < 0)
				throw new ArgumentOutOfRangeException("desiredNumberOfArguments", "may not be a negative value");

			var supportFunctionMatches = typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests)
				.GetMethods(BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance)
				.Where(m => m.Name.Equals(builtInFunctionToken.Content, StringComparison.OrdinalIgnoreCase));
			if (!supportFunctionMatches.Any())
				throw new NotSupportedException("Unsupported BuiltInFunctionToken content: " + builtInFunctionToken.Content);

			var idealMatch = supportFunctionMatches
				.Where(m => m.GetParameters().All(p => !p.IsOut && !p.ParameterType.IsByRef && p.ParameterType == typeof(object)))
				.FirstOrDefault(m => m.GetParameters().Length == desiredNumberOfArguments);
			if (idealMatch != null)
			{
				// If located a support function with the desired number of arguments (in the correct format; "in only" and type "object") then return a
				// BuiltInFunctionDetails with the name from that match (just in case the case of the function name varies - the match above is case insensitive)
				return new BuiltInFunctionDetails(builtInFunctionToken, idealMatch.Name, desiredNumberOfArguments, idealMatch.ReturnType);
			}
			// If a match was found for the name but not the desired number of arguments then return a result with null "desiredNumberOfArgumentsMatchedAgainst"
			// and "returnTypeIfKnown" values (it doesn't matter which of the names we select - for cases where they vary by case, which is not expected but not
			// impossible - so just return the first matched support function name)
			return new BuiltInFunctionDetails(builtInFunctionToken, supportFunctionMatches.First().Name, null, null);
		}

		private TranslatedStatementContentDetailsWithContentType TranslateCallExpressionSegment(
			NameToken target,
			IEnumerable<IToken> targetMemberAccessTokens,
			IEnumerable<Expression> arguments,
			CallSetItemExpressionSegment.ArgumentBracketPresenceOptions? zeroArgumentBracketsPresence,
			ScopeAccessInformation scopeAccessInformation,
			int indexInCallSet,
			bool targetIsKnownToBeBuiltInFunction)
		{
			if (target == null)
				throw new ArgumentNullException("target");
			if (targetMemberAccessTokens == null)
				throw new ArgumentNullException("targetMemberAccessTokens");
			if (arguments == null)
				throw new ArgumentNullException("arguments");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indexInCallSet < 0)
				throw new ArgumentOutOfRangeException("indexInCallSet");

			var targetMemberAccessTokensArray = targetMemberAccessTokens.ToArray();
			if (targetMemberAccessTokensArray.Any(t => t == null))
				throw new ArgumentException("Null reference encountered in targetMemberAccessTokens set");
			var argumentsArray = arguments.ToArray();
			if (argumentsArray.Any(a => a == null))
				throw new ArgumentException("Null reference encountered in arguments set");

			if (argumentsArray.Length == 0)
			{
				if (zeroArgumentBracketsPresence == null)
					throw new ArgumentException("zeroArgumentBracketsPresence may not be null if there are zero arguments");
				if ((zeroArgumentBracketsPresence .Value != CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent)
				&& (zeroArgumentBracketsPresence.Value != CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Present))
					throw new ArgumentException("Invalid zeroArgumentBracketsPresence value");
			}
			else if (zeroArgumentBracketsPresence != null)
				throw new ArgumentException("zeroArgumentBracketsPresence must be null if there are arguments present");

			// If this is part of a CallSetExpression and is not the first item then there is no point trying to analyse the origin of the targetName (check
			// its scope, etc..) since this should be something of the form "_.CALL(this, _outer, "F", _.ARGS.Val(0))" - there is nothing to be gained from trying
			// to guess whether it's a function or what variables were accessed since this has already been done. (It's still important to check for
			// undeclared variables referenced in the arguments but that is all handled later on). The same applies if the target is known to be
			// a built-in function (such as CDate).
			var targetName = _nameRewriter.GetMemberAccessTokenName(target);
			DeclaredReferenceDetails targetReferenceDetailsIfAvailable;
			CSharpName nameOfTargetContainerIfRequired;
			if (targetIsKnownToBeBuiltInFunction || (indexInCallSet > 0))
			{
				targetReferenceDetailsIfAvailable = null;
				nameOfTargetContainerIfRequired = null;
			}
			else
			{
				targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(target, _nameRewriter);
				nameOfTargetContainerIfRequired = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(
					target,
					_envRefName,
					_outerRefName,
					_nameRewriter
				);
			}

			// If there are no member access tokens then we have to consider whether this is a function call or property access, we can find this
			// out by looking into the scope access information. Note: Further down, we rely on function / property calls being identified at this
			// point for cases where there are no target member accessors (it means that if we get further down and there are no target member
			// accessors that it must not be a function or property call).
			// - The call semantics are different for a function call, if there is a method "F" in the outermost scope then something like
			//   "_.CALL(this, _outer, "F", args)" would be generated but if "F" isn't a function then "_.CALL(this, _outer.F, new string[], args)"
			//   would be
			if (targetMemberAccessTokensArray.Length == 0)
			{
				if (targetReferenceDetailsIfAvailable != null)
				{
					if ((targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Function)
					|| (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Property))
					{
						// 2014-04-11 DWR: This used to try to call the function directly (rather than through .CALL) which caused two problems;
						// firstly, VBScript will throw runtime exceptions for invalid argument counts (or allow the error to be swallowed if wrapped
						// in On Error Resume Next) whereas this code would not compile if the wrong number of arguments was specified (which is a
						// discrepancy or breaking change, depending upon how you look upon it). Secondly, it couldn't correctly handle updating
						// arguments that should be passed ByRef. So now this code path uses .CALL as the code further down does.
						// Note: This relies upon the extension method which allows .CALL to be executed with a target reference and single named
						// member accessor but no arguments
						var memberCallVariablesAccessed = new NonNullImmutableList<NameToken>();
						var memberCallContent = new StringBuilder();
						memberCallContent.AppendFormat(
							"{0}.CALL(this, {1}, \"{2}\"", // Pass "this" as the "context" argument
							_supportRefName.Name,
							(nameOfTargetContainerIfRequired == null) ? "this" : nameOfTargetContainerIfRequired.Name,
							targetName
						);
						if (argumentsArray.Any())
						{
							memberCallContent.Append(", ");
							memberCallContent.Append(_supportRefName.Name);
							memberCallContent.Append(".ARGS");
							for (var index = 0; index < argumentsArray.Length; index++)
							{
								var argumentContent = TranslateAsArgumentContent(
									argumentsArray[index],
									scopeAccessInformation,
									forceAllArgumentsToBeByVal: targetIsKnownToBeBuiltInFunction
								);
								memberCallContent.Append(argumentContent.TranslatedContent);
								memberCallVariablesAccessed = memberCallVariablesAccessed.AddRange(
									argumentContent.VariablesAccessed
								);
							}
						}
						memberCallContent.Append(")");
						return new TranslatedStatementContentDetailsWithContentType(
							memberCallContent.ToString(),
							ExpressionReturnTypeOptions.NotSpecified,
							memberCallVariablesAccessed
						);
					}
				}
			}

			// The "master" CALL method signature is
			//
			//   CALL(object target, IEnumerable<string> members, object[] arguments)
			//
			// (the arguments set is passed as an array as VBScript parameters are, by default, by-ref and so all of the arguments have to be
			// passed in this manner in case any of them need to be access in this manner).
			//
			// However, there are alternate signatures to try to make the most common calls easier to read - eg.
			//
			//   CALL(object)
			//   CALL(object, object[] arguments)
			//   CALL(object, string member1, object[] arguments)
			//   CALL(object, string member1, string member2, object[] arguments)
			//   ..
			//   CALL(object, string member1, string member2, string member3, string member4, string member5, object[] arguments)
			//
			// The maximum number of member access tokens (where only the targetMemberAccessTokens are considered, not the initial targetName
			// token) that may use one of these alternate signatures is stored in a constant in the IProvideVBScriptCompatFunctionality
			// static extension class that provides these convenience signatures.

			// Note: Even if there is only a single member access token (meaning call for "a" not "a.Test") and there are no arguments, we
			// still need to use the CALL method to account for any handling of default properties (eg. a statement "Test" may be a call to
			// a method named "Test" or "Test" may be an instance of a class which has a default parameter-less function, in which case the
			// default function will be executed by that statement.
			// 2014-04-10 DWR: This is not correct since if there are target member accessors and the target can be identified as a function
			// or property (according to the scope access information) then it would have been caught above. So if there are no target member
			// accessors or arguments then we can return a direct reference to the target here. Note that if a single member access token
			// constituted the entire statement then it would have to be forced through a .VAL call but that will also have been handled
			// before this point in the TryToGetShortCutStatementResponse call in the public Translate method (see notes in the
			// TryToGetShortCutStatementResponse method for more information about this).
			// 2015-05-31 DWR: Argh.. this was still not entirely correct :( We can ONLY take this shortcut if there are no arguments AND there
			// were no brackets around this absence-of-arguments, otherwise the brackets MAY have significance. Examples: If "a" is a number and
			// you try to access "a()" then you get a Type mismatch (it's not an array). If "a" is an array then "a()" throws a Subscript out of
			// range. If "a" is a VBScript class with a property "Name" then "a.Name()" will return the value, the same as "a.Name" (without
			// brackets), due to how VBScript describes classes internally. If "a" is an IDispatch reference then "a.Name()" will only work
			// if the target declares it has a method called "Name" (that will return a value for zero arguments), it may have a non-indexed
			// property called "Name" but it may not consider that applicable for a request for "a.Name()" because the brackets signify a
			// method, rather than property. So, if there are brackets then the CALL method must be used so that this logic can be
			// applied - only if there are no arguments and no brackets may the value be returned unwrapped.
			if (!targetMemberAccessTokensArray.Any() && !argumentsArray.Any() && (zeroArgumentBracketsPresence == CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent))
			{
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"{0}{1}",
						(nameOfTargetContainerIfRequired == null) ? "" : string.Format("{0}.", nameOfTargetContainerIfRequired.Name),
						targetName
					),
					ExpressionReturnTypeOptions.NotSpecified, // This could be anything so we have to report NotSpecified as the return type
					new NonNullImmutableList<NameToken>()
				);
			}

			var callExpressionContent = new StringBuilder();
			callExpressionContent.AppendFormat(
				"{0}.CALL(this, {1}{2}", // Pass "this" as the "context" argument
				_supportRefName.Name,
				(nameOfTargetContainerIfRequired == null) ? "" : string.Format("{0}.", nameOfTargetContainerIfRequired.Name),
				targetName
			);

			var ableToUseShorthandCallSignature = (targetMemberAccessTokensArray.Length <= IAccessValuesUsingVBScriptRules_Extensions.MaxNumberOfMemberAccessorBeforeArraysRequired);
			if (targetMemberAccessTokensArray.Length > 0)
			{
				callExpressionContent.Append(", ");
				if (!ableToUseShorthandCallSignature)
					callExpressionContent.Append(" new[] { ");
				for (var index = 0; index < targetMemberAccessTokensArray.Length; index++)
				{
					callExpressionContent.Append(
						targetMemberAccessTokensArray[index].Content.ToLiteral()
					);
					if (index < (targetMemberAccessTokensArray.Length - 1))
						callExpressionContent.Append(", ");
				}
				if (!ableToUseShorthandCallSignature)
					callExpressionContent.Append(" }");
			}

			var callExpressionVariablesAccessed = new NonNullImmutableList<NameToken>();
			if (argumentsArray.Length > 0)
			{
				callExpressionContent.Append(", ");
				var argumentProviderContent = TranslateAsArgumentProvider(argumentsArray, scopeAccessInformation, forceAllArgumentsToBeByVal: targetIsKnownToBeBuiltInFunction);
				callExpressionContent.Append(argumentProviderContent.TranslatedContent);
				callExpressionVariablesAccessed = callExpressionVariablesAccessed.AddRange(
					argumentProviderContent.VariablesAccessed
				);
			}
			else if (zeroArgumentBracketsPresence == CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Present)
			{
				callExpressionContent.AppendFormat(
					", {0}.ARGS.ForceBrackets()",
					_supportRefName.Name
				);
			}

			callExpressionContent.Append(")");
			return new TranslatedStatementContentDetailsWithContentType(
				callExpressionContent.ToString(),
				ExpressionReturnTypeOptions.NotSpecified, // This could be anything so we have to report NotSpecified as the return type
				callExpressionVariablesAccessed
			);
		}

		/// <summary>
		/// This generates the content that initialises a new IProvideCallArguments instance, based upon the specified argument values. This will throw
		/// an exception for null arguments or an argumentValues set containing any null references. It will never return null, it will raise an exception
		/// if unable to satisfy the request.
		/// </summary>
		public TranslatedStatementContentDetails TranslateAsArgumentProvider(
			IEnumerable<Expression> argumentValues,
			ScopeAccessInformation scopeAccessInformation,
			bool forceAllArgumentsToBeByVal)
		{
			if (argumentValues == null)
				throw new ArgumentNullException("argumentValues");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var variablesAccessed = new NonNullImmutableList<NameToken>();
			var argumentProviderContent = new StringBuilder();
			argumentProviderContent.Append(_supportRefName.Name);
			argumentProviderContent.Append(".ARGS");
			foreach (var argumentValue in argumentValues)
			{
				if (argumentValue == null)
					throw new ArgumentException("Null reference encountered in argumentValues set");

				var argumentContent = TranslateAsArgumentContent(argumentValue, scopeAccessInformation, forceAllArgumentsToBeByVal);
				argumentProviderContent.Append(argumentContent.TranslatedContent);
				variablesAccessed = variablesAccessed.AddRange(
					argumentContent.VariablesAccessed
				);
			}
			return new TranslatedStatementContentDetails(
				argumentProviderContent.ToString(),
				variablesAccessed
			);
		}

		/// <summary>
		/// If a call expression is for a support function and all of the arguments are of type object and none of them are out or ref arguments, and the number
		/// of arguments in a call expression matches the signature of the support function, then the support function can be called directly instead of going
		/// thorough a CALL execution. This makes the output code less verbose. If there is no matching signature then CALL must be used so that the argument
		/// mismatch becomes a runtime error, rather than compile time (since that's what VBScript does).
		/// </summary>
		private TranslatedStatementContentDetailsWithContentType TranslateAsDirectSupportFunctionCall(
			BuiltInFunctionDetails function,
			IEnumerable<Expression> argumentValues,
			ScopeAccessInformation scopeAccessInformation)
		{
			if (function == null)
				throw new ArgumentNullException("function");
			if (argumentValues == null)
				throw new ArgumentNullException("argumentValues");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var variablesAccessed = new NonNullImmutableList<NameToken>();
			var supportFunctionCallContent = new StringBuilder();
			supportFunctionCallContent.Append(_supportRefName.Name);
			supportFunctionCallContent.Append(".");
			supportFunctionCallContent.Append(function.SupportFunctionName);
			supportFunctionCallContent.Append("(");
			foreach (var indexedArgumentValue in argumentValues.Select((arg, index) => new { Argument = arg, Index = index }))
			{
				var argumentValue = indexedArgumentValue.Argument;
				if (argumentValue == null)
					throw new ArgumentException("Null reference encountered in argumentValues set");

				if (indexedArgumentValue.Index > 0)
					supportFunctionCallContent.Append(", ");

				// If this is a builtin function that is known to return a numeric type and it has a numeric constant as its argument(s), then we can emit
				// just the numeric value and drop any type information, since this will be thrown away when the number-returning function does its work
				// - eg. don't emit "CDBL((Int16)1)" since it's going to return a double so the Int16 type information is useless, "CDBL(1)" is no less
				// useful while being more succinct (and more natural when compared to the source it's being translated from)
				var argumentValueAsNumericConstant = TryToGetExpressionAsSingleNumericValueExpressionSegment(argumentValue);
				if ((argumentValueAsNumericConstant != null) && (function.Token is BuiltInFunctionToken) && ((BuiltInFunctionToken)function.Token).GuaranteedToReturnNumericContent)
				{
					supportFunctionCallContent.Append(argumentValueAsNumericConstant.Token.Value);
					continue;
				}

				var argumentContent = Translate(
					argumentValue,
					scopeAccessInformation,
					ExpressionReturnTypeOptions.NotSpecified
				);
				supportFunctionCallContent.Append(argumentContent.TranslatedContent);
				variablesAccessed = variablesAccessed.AddRange(
					argumentContent.VariablesAccessed
				);
			}
			supportFunctionCallContent.Append(")");
			ExpressionReturnTypeOptions supportFunctionReturnType;
			if (function.ReturnTypeIfKnown == null)
				supportFunctionReturnType = ExpressionReturnTypeOptions.NotSpecified;
			else if (function.ReturnTypeIfKnown == typeof(void))
				supportFunctionReturnType = ExpressionReturnTypeOptions.None;
			else if ((function.ReturnTypeIfKnown.IsValueType) || (function.ReturnTypeIfKnown == typeof(string)) || function.ReturnTypeIfKnown.IsArray)
				supportFunctionReturnType = ExpressionReturnTypeOptions.Value;
			else if (function.ReturnTypeIfKnown == typeof(object))
			{
				// If it's got "object" return type then it might be because it needs to be a string OR DBNull.Value, or it might genuinely be because
				// it needs to return a Reference type object. There's no way to know so we have to resort to NotSpecified.
				supportFunctionReturnType = ExpressionReturnTypeOptions.NotSpecified;
			}
			else
			{
				// If it's a non-object return type (that which is too vague to reason about) and it isn't a type that we know is Value type, then it
				// must be a Reference type.
				supportFunctionReturnType = ExpressionReturnTypeOptions.Reference;
			}
			return new TranslatedStatementContentDetailsWithContentType(
				supportFunctionCallContent.ToString(),
				supportFunctionReturnType,
				variablesAccessed
			);
		}

		private NumericValueExpressionSegment TryToGetExpressionAsSingleNumericValueExpressionSegment(Expression expression)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");

			if (expression.Segments.Count() != 1)
				return null;

			return expression.Segments.Single() as NumericValueExpressionSegment;
		}

		/// <summary>
		/// This generates the calls to IBuildCallArgumentProviders (such as .Val or .Ref) based upon argument content (eg. an argument value that
		/// is the result of calling another function can never be passed ByRef whereas an argument that is a simple variable reference always
		/// CAN be passed ByRef). Note that this does not have to consider whether the target function describes arguments as ByVal or ByRef,
		/// this is just about setting up the IBuildCallArgumentProviders data so that any ByRef arguments CAN be updated on the caller
		/// where required.
		/// </summary>
		private TranslatedStatementContentDetails TranslateAsArgumentContent(
			Expression argumentValue,
			ScopeAccessInformation scopeAccessInformation,
			bool forceAllArgumentsToBeByVal)
		{
			if (argumentValue == null)
				throw new ArgumentNullException("argumentValue");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var isConfirmedToBeByVal = forceAllArgumentsToBeByVal || ArgumentWouldBePassedByValBasedUponItsContent(argumentValue, scopeAccessInformation);
			if (isConfirmedToBeByVal)
			{
				var translatedCallExpressionByValArgumentContent = Translate(
					argumentValue,
					scopeAccessInformation,
					ExpressionReturnTypeOptions.NotSpecified
				);
				return new TranslatedStatementContentDetails(
					string.Format(
						".Val({0})",
						translatedCallExpressionByValArgumentContent.TranslatedContent
					),
					translatedCallExpressionByValArgumentContent.VariablesAccessed
				);
			}

			// If we've got this far then then argumentValue expression has only a single segment. If it is a CallExpressionSegment or a CallSetExpression
			// then there won't be any nested member accessors (such as "a.Name" or "a(0).Name" since they would have been caught as a ByVal situation
			// above). On the other hand, there shouldn't be any other type of expression segment that could get this far either! (Access of a variable
			// "a" is represented by a CallExpressionSegment with a single member accessor token and zero arguments). The only easy out we have at this
			// point is if we do indeed have a CallExpressionSegment with a single member accessor token and zero arguments since that will definitely
			// be passed ByRef. Otherwise there are arguments to consider which may be arguments on default functions or properties (in which case it
			// will be ByVal) or array indices (in which case it will be ByRef).
			if ((argumentValue.Segments.Single() is CallExpressionSegment) && !((CallExpressionSegment)argumentValue.Segments.Single()).Arguments.Any())
			{
				var translatedCallExpressionByRefArgumentContent = Translate(
					argumentValue,
					scopeAccessInformation,
					ExpressionReturnTypeOptions.NotSpecified
				);
				return new TranslatedStatementContentDetails(
					string.Format(
						".Ref({0}, {1} => {{ {0} = {1}; }})",
						translatedCallExpressionByRefArgumentContent.TranslatedContent,
						_tempNameGenerator(new CSharpName("v"), scopeAccessInformation).Name
					),
					translatedCallExpressionByRefArgumentContent.VariablesAccessed
				);
			}

			// The final scenario is a CallExpressionSegment with a single member accessor and one or more arguments or a CallSetExpressionSegment
			// where the first segment has a single member accessor and one or more arguments and the subsequent segments all have zero member
			// accessors but one or more arguments. It's impossible to know at this point whether those "arguments" are array accesses (which will
			// be passed ByRef) or default function or property calls (which will be passed ByVal).
			TranslatedStatementContentDetails possibleByRefTarget;
			IEnumerable<IEnumerable<Expression>> possibleByRefArgumentSets;
			var possibleByRefCallExpressionSegment = argumentValue.Segments.Single() as CallExpressionSegment;
			if (possibleByRefCallExpressionSegment != null)
			{
				// Note: Know that there is at least one argument, so possibleByRefCallExpressionSegment's ZeroArgumentBracketsPresence value will
				// be null, but we need a non-null value since we'll be creating a new CallExpressionSegment that targets the by-ref alias but that
				// strips off the arguments for now. If the value we used was "Present" then we would risk introducing additional CALL logic that
				// we don't want, so specify it as "Absent".
				if (possibleByRefCallExpressionSegment.MemberAccessTokens.Count() > 2)
					throw new NotSupportedException("Unexpected argumentValue content - didn't expect a CallExpressionSegment with multiple MemberAccessTokens at this point");
				if (!possibleByRefCallExpressionSegment.MemberAccessTokens.Any())
					throw new NotSupportedException("Unexpected argumentValue content - didn't expect a CallExpressionSegment without any arguments at this point");
				possibleByRefTarget = Translate(
					new CallExpressionSegment(
						possibleByRefCallExpressionSegment.MemberAccessTokens,
						new Expression[0],
						CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent
					),
					scopeAccessInformation
				);
				possibleByRefArgumentSets = new[] { possibleByRefCallExpressionSegment.Arguments };
			}
			else
			{
				var possibleByRefCallSetExpressionSegment = argumentValue.Segments.Single() as CallSetExpressionSegment;
				if (possibleByRefCallSetExpressionSegment != null)
				{
					if (possibleByRefCallSetExpressionSegment.CallExpressionSegments.First().MemberAccessTokens.Count() > 2)
						throw new NotSupportedException("Unexpected argumentValue content - didn't expect a CallSetExpressionSegment with multiple MemberAccessTokens in its first CallExpressionSegment at this point");
					if (possibleByRefCallSetExpressionSegment.CallExpressionSegments.Skip(1).Any(s => s.MemberAccessTokens.Any()))
						throw new NotSupportedException("Unexpected argumentValue content - didn't expect a CallSetExpressionSegment with subsequent CallExpressionSegments that have MemberAccessTokens at this point");

					// Note: Specify CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent for the same reason as explain in the logic above
					possibleByRefTarget = Translate(
						new CallExpressionSegment(
							possibleByRefCallSetExpressionSegment.CallExpressionSegments.First().MemberAccessTokens,
							new Expression[0],
							CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent
						),
						scopeAccessInformation
					);
					possibleByRefArgumentSets = possibleByRefCallSetExpressionSegment.CallExpressionSegments.Select(s => s.Arguments);
				}
				else
					throw new NotSupportedException("Unexpected argumentValue content, unable to translate");
			}

			// For the "RefIfArray" call to determine the correct behaviour at runtime, we need to pass the initial target to the method and then
			// each set of arguments so that it can check at each point whether the arguments are for a default function/property call or whether
			// they are for array access (as soon as a function or property call is made, the value must be passed ByVal). This means that a call
			// "a(0, 1)(2)" is effectively passed as "RefIfArray(a, (0, 1), (2))". If it only checked whether the (2) argument was for an array
			// access or a function/property call then it would ignore whether the (0, 1) arguments were for array access or function/property.
			// - Note: This RefIfArray call relies upon the extension method that takes a param array instead of an IEnumerable
			var translatedContentForPossibleByRefArgumentSets = possibleByRefArgumentSets
				.Select(args => TranslateAsArgumentProvider(args, scopeAccessInformation, forceAllArgumentsToBeByVal: false));
			return new TranslatedStatementContentDetails(
				string.Format(
					".RefIfArray({0}, {1})",
					possibleByRefTarget.TranslatedContent,
					string.Join(
						", ",
						translatedContentForPossibleByRefArgumentSets.Select(content => content.TranslatedContent)
					)
				),
				possibleByRefTarget.VariablesAccessed.AddRange(
					translatedContentForPossibleByRefArgumentSets.SelectMany(content => content.VariablesAccessed)
				)
			);
		}

		/// <summary>
		/// It is possible to determine by analysing the content of a function/property argument whether it will be passed ByVal even if the
		/// argument on the target function/property states ByRef. This will obviously be the case for constants but may also be the case for
		/// variables, depending upon how they are accessed - if an argument is "a" and "a" is a variable then it is not possible to determine
		/// from its content that it should be passed ByRef, but if the argument is "a.Name" then it can be seen that it must be passed ByVal
		/// since VBScript will not pass the "Name" property as ByRef to the function. The same goes for function/property calls such as
		/// "a.GetValue(0)" but not for default function/property calls such as "a(0)" since this may be a default member access or it
		/// may be an array access - in the case of an array element, that must be passed ByRef and so it is not safe to say that "a(0)"
		/// may definitely be passed ByVal.
		/// </summary>
		public bool ArgumentWouldBePassedByValBasedUponItsContent(Expression argumentValue, ScopeAccessInformation scopeAccessInformation)
		{
			if (argumentValue == null)
				throw new ArgumentNullException("argumentValue");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			if (argumentValue.Segments.Count() > 1)
			{
				// If there are multiple segments here then it must be ByVal (it's fairly difficult to actually not be ByVal - it basically boils
				// down to being a simple variable reference - eg. "a" - or one or more array accesses - eg. "a(0)" or "a(0, 1)" or "a(0)(1)"
				// though differentiating between the case where "a(0)" is an array accessor and where it's a default function/property
				// access is where some of the difficulty comes in). All of the below will try to decide what to do.
				return true;
			}

			// If there is a single segment that is a constant then it must be ByVal, same if it's a bracketed segment (VBScript treats values
			// passed wrapped in extra brackets as being ByVal), same if it's a NewInstanceExpressionSegment (there is no point allowing the
			// reference to be changed since nothing has a reference to the new instance being passed in)
			var singleSegment = argumentValue.Segments.First();
			if ((singleSegment is DateValueExpressionSegment)
			|| (singleSegment is NumericValueExpressionSegment)
			|| (singleSegment is StringValueExpressionSegment)
			|| (singleSegment is BuiltInValueExpressionSegment)
			|| (singleSegment is NewInstanceExpressionSegment)
			|| (singleSegment is RuntimeErrorExpressionSegment))
				return true;

			if (singleSegment is BracketedExpressionSegment)
			{
				// If this argument content is bracketed then it must be passed ByVal (a VBScript peculiarity) but if the brackets are present
				// for that purpose only then they needn't be included in the final output, so if there is a single term that has been wrapped
				// in brackets then unwrap the term (overwrite the argumentValue reference so that this change is reflected in the rendering
				// call further down).
				while ((singleSegment is BracketedExpressionSegment) && (((BracketedExpressionSegment)singleSegment).Segments.Count() == 1))
					singleSegment = ((BracketedExpressionSegment)singleSegment).Segments.First();
				argumentValue = new Expression(new[] { singleSegment });
				return true;
			}

			// If this is either a single CallExpressionSegment or a CallSetExpressionSegment then there are some conditions that will mean that
			// it should definitely be passed ByVal. If there are any nested member accessors (eg. "a.Name" or "a(0).Name") then pass ByVal. If
			// the first call segment is to a built-in function then pass ByVal
			bool isConfirmedToBeByVal;
			CallSetItemExpressionSegment initialCallSetItemExpressionSegmentToCheckIfAny;
			var callExpressionSegment = singleSegment as CallExpressionSegment;
			if (callExpressionSegment != null)
			{
				initialCallSetItemExpressionSegmentToCheckIfAny = callExpressionSegment;
				isConfirmedToBeByVal = false; // Note: This may be overridden below since we have a non-null initialCallSetItemExpressionSegmentToCheckIfAny
			}
			else
			{
				var callSetExpressionSegment = singleSegment as CallSetExpressionSegment;
				if (callSetExpressionSegment != null)
				{
					// The first call expression segment must have at least one member access tokens (otherwise there would be no target) but
					// if there are multiple (checked for further down) or if any of the subsequent segments have ANY accessors) then it's ByVal
					// (a CallSetExpression would have segments without any member accessors if it was describing jagged array access - eg.
					// "a(0)(1)" - or if the same source code was accessing default properties or members).
					if (callSetExpressionSegment.CallExpressionSegments.Skip(1).Any(s => s.MemberAccessTokens.Any()))
					{
						isConfirmedToBeByVal = true;
						initialCallSetItemExpressionSegmentToCheckIfAny = null; // No point doing any more checks so set this to null
					}
					else
					{
						initialCallSetItemExpressionSegmentToCheckIfAny = callSetExpressionSegment.CallExpressionSegments.First();
						isConfirmedToBeByVal = false; // Note: This may be overridden below since we have a non-null initialCallSetItemExpressionSegmentToCheckIfAny
					}
				}
				else
				{
					initialCallSetItemExpressionSegmentToCheckIfAny = null;
					isConfirmedToBeByVal = false;
				}
			}
			if (initialCallSetItemExpressionSegmentToCheckIfAny != null)
			{
				// If this is a call with multiple member accessors (indicating property access - eg. "a.Name") then it's passed ByVal. If it's
				// a built-in function then it's passed ByVal. If it's a known function within the current scope then it's passed ByVal.
				// - Check for multiple member accessor or built-in function access first, cos it's easy..
				if ((initialCallSetItemExpressionSegmentToCheckIfAny.MemberAccessTokens.Count() > 1)
				|| (initialCallSetItemExpressionSegmentToCheckIfAny.MemberAccessTokens.First() is BuiltInFunctionToken))
					isConfirmedToBeByVal = true;
				else
				{
					// .. then check for a known function
					var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(
						(NameToken)initialCallSetItemExpressionSegmentToCheckIfAny.MemberAccessTokens.First(), // TODO: Are we sure this is always going to be a NameToken??
						_nameRewriter
					);
					if (targetReferenceDetailsIfAvailable == null)
					{
						// If this is an undeclared reference then it will be implicitly declared later as a variable and so will be elligible
						// to be passed ByRef
						isConfirmedToBeByVal = false;
					}
					else
					{
						isConfirmedToBeByVal = (
							(targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Constant) ||
							(targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Function) ||
							(targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Property)
						);
					}
				}
			}
			return isConfirmedToBeByVal;
		}

		private TranslatedStatementContentDetailsWithContentType Translate(CallSetExpressionSegment callSetExpressionSegment, ScopeAccessInformation scopeAccessInformation)
		{
			if (callSetExpressionSegment == null)
				throw new ArgumentNullException("callSetExpressionSegment");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var content = "";
			var variablesAccessed = new NonNullImmutableList<NameToken>();
			var numberOfCallExpressions = callSetExpressionSegment.CallExpressionSegments.Count(); // This will always be at least two (see notes in CallSetExpressionSegment)
			for (var index = 0; index < numberOfCallExpressions; index++)
			{
				var callSetItemExpression = callSetExpressionSegment.CallExpressionSegments.ElementAt(index);
				TranslatedStatementContentDetailsWithContentType translatedContent;
				if (index == 0)
				{
					// The first segment in a CallSetExpressionSegment will always have the data required to represent a CallExpressionSegment
					// (the only difference in the types is that a CallSetItemExpressionSegment may have zero Member Access Tokens whereas the
					// CallExpressionSegment must always have at least one - however, the first segment in a CallSetExpressionSegment will
					// always have at least one as well)
					translatedContent = Translate(
						new CallExpressionSegment(
							callSetItemExpression.MemberAccessTokens,
							callSetItemExpression.Arguments,
							callSetItemExpression.ZeroArgumentBracketsPresence
						),
						scopeAccessInformation
					);
				}
				else
				{
					translatedContent = TranslateCallExpressionSegment(
						new DoNotRenameNameToken(content, ((IExpressionSegment)callSetExpressionSegment).AllTokens.First().LineIndex),
						callSetItemExpression.MemberAccessTokens,
						callSetItemExpression.Arguments,
						callSetItemExpression.ZeroArgumentBracketsPresence,
						scopeAccessInformation,
						index,
						targetIsKnownToBeBuiltInFunction: false
					);
				}

				// Overwrite any previous string content since it effectively gets passed through as an accumulator
				content = translatedContent.TranslatedContent;
				variablesAccessed = variablesAccessed.AddRange(translatedContent.VariablesAccessed);
			}
			return new TranslatedStatementContentDetailsWithContentType(
				content,
				ExpressionReturnTypeOptions.NotSpecified, // This could be anything so we have to report NotSpecified as the return type
				variablesAccessed
			);
		}

		private TranslatedStatementContentDetailsWithContentType Translate(
			NewInstanceExpressionSegment newInstanceExpressionSegment,
			VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions scopeLocation)
		{
			if (newInstanceExpressionSegment == null)
				throw new ArgumentNullException("newInstanceExpressionSegment");
			if (!Enum.IsDefined(typeof(VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions), scopeLocation))
				throw new ArgumentOutOfRangeException("scopeLocation");

			// Deal with the special case of "new RegExp" - this is a built-in class and not a class that will be translated (it is the
			// responsibility of the runtime support class to provide it)
			if (newInstanceExpressionSegment.ClassName.Content.Equals("RegExp", StringComparison.OrdinalIgnoreCase))
			{
				return new TranslatedStatementContentDetailsWithContentType(
					string.Format(
						"{0}.NEWREGEXP()",
						_supportRefName.Name
					),
					ExpressionReturnTypeOptions.Reference,
					new NonNullImmutableList<NameToken>()
				);
			}

			// Send the new instance to the NEW method so that it can be tracked and forcibly disposed of after ther request
			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"{0}.NEW(new {1}({0}, {2}, {3}))",
					_supportRefName.Name,
					_nameRewriter.GetMemberAccessTokenName(newInstanceExpressionSegment.ClassName),
					_envRefName.Name,
					_outerRefName.Name
				),
				ExpressionReturnTypeOptions.Reference,
				new NonNullImmutableList<NameToken>()
			);
		}

		private TranslatedStatementContentDetailsWithContentType Translate(RuntimeErrorExpressionSegment runtimeErrorExpressionSegment)
		{
			if (runtimeErrorExpressionSegment == null)
				throw new ArgumentNullException("runtimeErrorExpressionSegment");

			// This expression segment is generated when there is VBScript that is known to cause a runtime error - an exception needs to be thrown when
			// the translated code is run, but not at compile time (since runtime errors can be trapped with ON ERROR RESUME NEXT, but not if they cause
			// the translation process to blow up!)
			// - eg. "WScript.Echo 1()" will result in a type mismatch since the numeric constant can not be called like a function
			// Note: Translated programs will include a "using" for the namespace containing the VBScript-specific exceptions, which are what are most
			// commonly expected here. If the required exception is within the same namespace as SpecificVBScriptException then it need only be specified
			// by name, otherwise its "FullName" will be required (which includes its namespace).
			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"{0}.RAISEERROR(new {1}({2}))",
					_supportRefName.Name,
					(runtimeErrorExpressionSegment.ExceptionType.Namespace == typeof(SpecificVBScriptException).Namespace)
						? runtimeErrorExpressionSegment.ExceptionType.Name
						: runtimeErrorExpressionSegment.ExceptionType.FullName,
					runtimeErrorExpressionSegment.Message.ToLiteral()
				),
				ExpressionReturnTypeOptions.Reference,
				new NonNullImmutableList<NameToken>()
			);
		}

		private string GetSupportFunctionName(OperatorToken operatorToken)
		{
			if (operatorToken == null)
				throw new ArgumentNullException("operatorToken");

			switch (operatorToken.Content.ToUpper())
			{
				// Arithmetic operators
				case "^": return "POW";
				case "/": return "DIV";
				case "*": return "MULT";
				case "\\": return "INTDIV";
				case "MOD": return "MOD";
				case "+": return "ADD";
				case "-": return "SUBT";

				// String concatenation
				case "&": return "CONCAT";

				// Logical operators
				case "NOT": return "NOT";
				case "AND": return "AND";
				case "OR": return "OR";
				case "XOR": return "XOR";

				// Comparison operators
				case "=": return "EQ";
				case "<>": return "NOTEQ";
				case "<": return "LT";
				case ">": return "GT";
				case "<=": return "LTE";
				case ">=": return "GTE";
				case "IS": return "IS";
				case "EQV": return "EQV";
				case "IMP": return "IMP";

				// Error
				default:
					throw new NotSupportedException("Unsupported OperatorToken content: " + operatorToken.Content);
			}
		}

		private string ApplyReturnTypeGuarantee(string translatedContent, ExpressionReturnTypeOptions contentType, ExpressionReturnTypeOptions requiredReturnType, int lineIndex)
		{
			if (string.IsNullOrWhiteSpace(translatedContent))
				throw new ArgumentException("Null/blank translatedContent specified");
			if (lineIndex < 0)
				throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

			switch (requiredReturnType)
			{
				case ExpressionReturnTypeOptions.Boolean:
					return string.Format(
						"{0}.IF({1})",
						_supportRefName.Name,
						translatedContent
					);

				case ExpressionReturnTypeOptions.NotSpecified:
				case ExpressionReturnTypeOptions.None:
					return translatedContent;

				case ExpressionReturnTypeOptions.Reference:
					// If we know that this returns a value type then we can tell at this point that it's not going to work. If it returns a Reference
					// type then we're golden. If contentType is NotSpecified then we need to pass it through the OBJ method so that a runtime exception
					// is raised if the expression is NOT a reference type, in order to be consistent with VBScript's behaviour. (Previously, this would
					// throw an exception at "translation time" - the runtime of the translator, as opposed to the runtime of the generated C# - that
					// would indicate that the content was invalid for a Reference result if contentType was Boolean or Value, but this is inconsistent
					// with VBScript, which would throw an exception at runtime - equivalent to the generated C#'s runtime. Now a log warning is
					// recorded, which is hopefully a reasonable compromise).
					if (contentType == ExpressionReturnTypeOptions.Reference)
						return translatedContent;
					if ((contentType == ExpressionReturnTypeOptions.Boolean)
					|| (contentType == ExpressionReturnTypeOptions.None)
					|| (contentType == ExpressionReturnTypeOptions.Value))
						_logger.Warning("Request for an object reference at line " + (lineIndex + 1) + " but data type is " + contentType);
					return string.Format(
						"{0}.OBJ({1})",
						_supportRefName.Name,
						translatedContent
					);

				case ExpressionReturnTypeOptions.Value:
					if ((contentType == ExpressionReturnTypeOptions.Boolean) || contentType == ExpressionReturnTypeOptions.Value)
						return translatedContent;
					return string.Format(
						"{0}.VAL({1})",
						_supportRefName.Name,
						translatedContent
					);

				default:
					throw new NotSupportedException("Unsupported requiredReturnType value: " + requiredReturnType);
			}
		}

		private TranslatedStatementContentDetails TryToGetShortCutStatementResponse(
			Expression expression,
			ScopeAccessInformation scopeAccessInformation,
			ExpressionReturnTypeOptions returnRequirements)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			// There are some special rules for Statements (expressions where the return type is None); if "o" is an object reference, then a statement "o" will have value-
			// type logic applied to it, so it will require a parameter-less default function or property otherwise an "Object doesn't support this property or method"
			// error will be raised. However, if "o" is a function which returns an object reference, the value-type logic will not be applied to it. If "o" has a property
			// "Child" which is an object reference, then "o.Child" will also not have value-type access logic applied to it. If "o" is an array, then a statement "o(0)"
			// will result in a "Type mismatch" error; "o(0).Name" would not, however (assuming that the first element of the "o" array was an object with a property
			// called "Name").
			// TODO: The logic applied here does not cover all error cases; if the single NameToken is a value (eg. 1 or "one") then an "Expected Statement" error is
			// raised. If the NameToken is a non-object value (eg. i where i is null or 1 or "one") then a "Type mistmatch" error is raised. If the NameToken is an
			// object reference without a default member then an "Object doesn't support this property or method" error is raised unless the object reference is to
			// Nothing, in which case a "Type mismatch error" is raised. If the NamToken is an object reference WITH a default member then no error occurs (the
			// default member is accessed / executed).

			// The specific case that we want to address with this "shortcut" avenue is where we have a statement (so return type is None) with a single call expression
			// segment with a single NameToken, which does not indicate a function or property. If these conditions are met then we can avoid all of the rest of the
			// translation process. Note: We can't do this if return type is anything other than None since in C# it's not valid to have a statement that is only an
			// instance of a class (if it's wrapped in a call to OBJ or VAL then it's ok since it's a method call, but that would be handled by the standard
			// translation process).
			if (returnRequirements != ExpressionReturnTypeOptions.None)
				return null;

			if (expression.Segments.Take(2).Count() > 1)
				return null;

			var onlyExpressionSegmentAsCallExpression = expression.Segments.Single() as CallExpressionSegment;
			if (onlyExpressionSegmentAsCallExpression == null)
				return null;

			// If there are multiple member accessor tokens, arguments or if there were brackets following the single member accessor then this logic doesn't apply
			if ((onlyExpressionSegmentAsCallExpression.MemberAccessTokens.Take(2).Count() > 1)
			|| onlyExpressionSegmentAsCallExpression.Arguments.Any()
			|| onlyExpressionSegmentAsCallExpression.ZeroArgumentBracketsPresence == CallExpressionSegment.ArgumentBracketPresenceOptions.Present)
				return null;

			var onlyMemberAccessTokenAsName = onlyExpressionSegmentAsCallExpression.MemberAccessTokens.Single() as NameToken;
			if (onlyMemberAccessTokenAsName == null)
				return null;

			// If this is a function of property then we can't consider it for this shortcut
			var rewrittenName = _nameRewriter.GetMemberAccessTokenName(onlyMemberAccessTokenAsName);
			var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(onlyMemberAccessTokenAsName, _nameRewriter);
			if ((targetReferenceDetailsIfAvailable == null) || (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.ExternalDependency))
				rewrittenName = _envRefName.Name + "." + rewrittenName;
			else if (targetReferenceDetailsIfAvailable != null)
			{
				if ((targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Function)
				|| (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Property))
					return null;

				if (targetReferenceDetailsIfAvailable.ScopeLocation == VBScriptTranslator.LegacyParser.CodeBlocks.Basic.ScopeLocationOptions.OutermostScope)
					rewrittenName = _outerRefName.Name + "." + rewrittenName;
			}

			// In fact, if we know there's only a single non-locally-scoped-function-or-property NameToken that needs to return a value type, we can just
			// return now. We can't do this if the return type is NotSpecified since in C# it's not valid to have a statement that is only an instance
			// of a class (but if it's wrapped in a call to OBJ or VAL then it's ok). This logic can only be applied to non-value-returning Statements,
			// Expressions that return values could exist as just "o" since that WOULD be valid C#.
			return new TranslatedStatementContentDetails(
				string.Format(
					"{0}.VAL({1})",
					_supportRefName.Name,
					rewrittenName
				),
				new NonNullImmutableList<NameToken>(new[] { onlyMemberAccessTokenAsName })
			);
		}

		/// <summary>
		/// It's very common in VBScript code to get runs of string concatenations, so rather then making the translated code longer than it needs to be, by limiting
		/// the number of arguments takens by the CONCAT method to two (like the other operators - except for NOT, which takes only one argument), the CONCAT method
		/// may also take more than two arguments if it would make no difference to the enforcing of operator precedence. This methods tries to identify cases where
		/// that might be applicable and performs a "special mode translation".
		/// </summary>
		private TranslatedStatementContentDetails TryToGetConcatFlattenedSpecialCaseResponse(
			Expression expression,
			ScopeAccessInformation scopeAccessInformation,
			ExpressionReturnTypeOptions returnRequirements)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
				throw new ArgumentOutOfRangeException("returnRequirements");

			// If there are three or less expression segments after any possible concat-flattening then this is not a special case, it can be handled
			// by the common flow for one, two or three segments
			var expressionSegmentsArray = ConcatFlattener.Flatten(expression).Segments.ToArray();
			if (expressionSegmentsArray.Length <= 3)
				return null;

			// The even segments must be values and the odd segments must all be concat operators
			var evenSegments = expressionSegmentsArray.Where((segment, index) => (index % 2) == 0);
			if (evenSegments.Any(s => s is OperationExpressionSegment))
				return null;
			var oddSegments = expressionSegmentsArray.Where((segment, index) => (index % 2) == 1);
			if (!oddSegments.All(s => s is OperationExpressionSegment) || oddSegments.OfType<OperationExpressionSegment>().Any(s => s.Token.Content != "&"))
				return null;

			var translatedNonOperatorSegments = evenSegments.Select(segment => TranslateNonOperatorSegment(segment, scopeAccessInformation));
			return new TranslatedStatementContentDetails(
				ApplyReturnTypeGuarantee(
					string.Format(
						"{0}.{1}({2})",
						_supportRefName.Name,
						GetSupportFunctionName(((OperationExpressionSegment)oddSegments.First()).Token),
						string.Join(", ", translatedNonOperatorSegments.Select(c => c.TranslatedContent))
					),
					ExpressionReturnTypeOptions.Value, // All VBScript operators return numeric (or boolean, which are also numeric in VBScript) values
					returnRequirements,
					expressionSegmentsArray[0].AllTokens.First().LineIndex
				),
				translatedNonOperatorSegments.SelectMany(c => c.VariablesAccessed).ToNonNullImmutableList()
			);
		}

		private class TranslatedStatementContentDetailsWithContentType : TranslatedStatementContentDetails
		{
			public TranslatedStatementContentDetailsWithContentType(
				string content,
				ExpressionReturnTypeOptions contentType,
				NonNullImmutableList<NameToken> variablesAccessed) : base(content, variablesAccessed)
			{
				if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), contentType))
					throw new ArgumentOutOfRangeException("contentType");

				ContentType = contentType;
			}

			/// <summary>
			/// This will never be null
			/// </summary>
			public ExpressionReturnTypeOptions ContentType { get; private set; }
		}

		private class BuiltInFunctionDetails
		{
			public BuiltInFunctionDetails(IToken token, string supportFunctionName, int? desiredNumberOfArgumentsMatchedAgainst, Type returnTypeIfKnown)
			{
				if (token == null)
					throw new ArgumentNullException("token");
				if (string.IsNullOrWhiteSpace(supportFunctionName))
					throw new ArgumentException("Null/blank name specified");
				if (((desiredNumberOfArgumentsMatchedAgainst == null) && (returnTypeIfKnown != null))
				|| ((desiredNumberOfArgumentsMatchedAgainst != null) && (returnTypeIfKnown == null)))
					throw new ArgumentException("If one of desiredNumberOfArgumentsMatchedAgainst and returnTypeIfKnown is null then they both must be and vice versa");

				Token = token;
				SupportFunctionName = supportFunctionName;
				DesiredNumberOfArgumentsMatchedAgainst = desiredNumberOfArgumentsMatchedAgainst;
				ReturnTypeIfKnown = returnTypeIfKnown;
			}

			/// <summary>
			/// This will never be null
			/// </summary>
			public IToken Token { get; private set; }

			/// <summary>
			/// This will never be null or blank
			/// </summary>
			public string SupportFunctionName { get; private set; }

			/// <summary>
			/// If details of this function were requested with a desired number of parameters to match, then this will be that number if it was possible
			/// to location a support function with the requested name and that number of parameters so long as every parameter was "in only" (not "out"
			/// and not "ref") and of type object. If a function with the requested name was available but not with the desired number of parameters,
			/// this will be null.
			/// </summary>
			public int? DesiredNumberOfArgumentsMatchedAgainst { get; private set; }

			/// <summary>
			/// This will be null if DesiredNumberOfArgumentsMatchedAgainst is null and non-null if not since it relies upon the same criteria when trying
			/// to identify a target support function.
			/// </summary>
			public Type ReturnTypeIfKnown { get; set; }
		}
	}
}
