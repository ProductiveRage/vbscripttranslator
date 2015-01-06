using CSharpSupport;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using LegacyParser = VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
    // TODO: Ensure the nameRewriter isn't being used anywhere it shouldn't - it shouldn't rename methods of COM components, for example
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
            bool mustConvertLeftValueToNumber, mustConvertRightValueToNumber, mustConvertLeftValueToString, mustConvertRightValueToString;
            if (operatorSegmentWithIndex.Segment.Token is ComparisonOperatorToken)
            {
                // If the operator segment is a ComparisonOperatorToken (not a LogicalOperatorToken and not any other type of OperatorToken, such
                // as an arithmetic operation or string concatenation) then there are special rules to apply if one side of the operation is known
                // to be a constant at compile time - namely, the other side must be parseable as a numeric value otherwise a "Type mismatch" will
                // be raised. If both sides of the operation are variables and one of them is a numeric value and the other a word, then the
                // comparison will return false but it won't raise an error. There is similar logic for string constants (but no equivalent
                // for boolean constants).
                // See http://blogs.msdn.com/b/ericlippert/archive/2004/07/30/202432.aspx for details about "hard types" that pertain to this
                if ((segmentLeft is NumericValueExpressionSegment) && (segmentRight is NumericValueExpressionSegment))
                {
                    // If both sides of an operation are numeric constants, then the comparison will be simple and not require any interfering in
                    // terms of casting values at this point
                    mustConvertLeftValueToNumber = false;
                    mustConvertRightValueToNumber = false;
                    mustConvertLeftValueToString = false;
                    mustConvertRightValueToString = false;
                }
                else if ((segmentLeft is NumericValueExpressionSegment) || (segmentRight is NumericValueExpressionSegment))
                {
                    // However, if one side of an operation is a numeric constant, then the other side must be parseable as a number - otherwise
                    // a "Type mismatch" error should be raised; the statement "IF ("aa" > 0) THEN" will error, for example
                    mustConvertLeftValueToNumber = !(segmentLeft is NumericValueExpressionSegment);
                    mustConvertRightValueToNumber = !(segmentRight is NumericValueExpressionSegment);
                    mustConvertLeftValueToString = false;
                    mustConvertRightValueToString = false;
                }
                else if ((segmentLeft is StringValueExpressionSegment) && (segmentRight is StringValueExpressionSegment))
                {
                    // If both sides of an operation are numeric constants, then the comparison will be simple and not require any interfering in
                    // terms of casting values at this point
                    mustConvertLeftValueToNumber = false;
                    mustConvertRightValueToNumber = false;
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
                    mustConvertLeftValueToString = !(segmentLeft is StringValueExpressionSegment);
                    mustConvertRightValueToString = !(segmentRight is StringValueExpressionSegment);
                }
                else
                {
                    // If neither side is a numeric constant then there is nothing to worry about (any complexities of how values should be
                    // compared can be left up to the runtime comparison method implementation)
                    mustConvertLeftValueToNumber = false;
                    mustConvertRightValueToNumber = false;
                    mustConvertLeftValueToString = false;
                    mustConvertRightValueToString = false;
                }
            }
            else
            {
                mustConvertLeftValueToNumber = false;
                mustConvertRightValueToNumber = false;
                mustConvertLeftValueToString = false;
                mustConvertRightValueToString = false;
            }
            var resultLeft = TranslateNonOperatorSegment(segmentLeft, scopeAccessInformation);
            if (mustConvertLeftValueToNumber && mustConvertLeftValueToString)
                throw new Exception("Something went wrong in the processing, both mustConvertLeftValueToNumber and mustConvertLeftValueToString are set for a comparison");
            if (mustConvertLeftValueToNumber)
                resultLeft = WrapTranslatedResultInNumericCast(resultLeft);
            else if (mustConvertLeftValueToString)
                resultLeft = WrapTranslatedResultInStringConversion(resultLeft);
            var resultRight = TranslateNonOperatorSegment(segmentRight, scopeAccessInformation);
            if (mustConvertRightValueToNumber && mustConvertRightValueToString)
                throw new Exception("Something went wrong in the processing, both mustConvertRightValueToNumber and mustConvertRightValueToString are set for a comparison");
            if (mustConvertRightValueToNumber)
                resultRight = WrapTranslatedResultInNumericCast(resultRight);
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

            return new TranslatedStatementContentDetailsWithContentType(
                string.Format(
                    "{0}.NUM({1})",
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
                    "{0}.STR({1})",
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
                    numericValueSegment.Token.Content,
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
			if (builtInValueExpressionSegment.Token.Content.Equals("nothing", StringComparison.InvariantCultureIgnoreCase))
			{
                return new TranslatedStatementContentDetailsWithContentType(
                    string.Format(
						"{0}.Constants.Nothing",
						_supportRefName.Name
					),
					ExpressionReturnTypeOptions.Reference,
                    new NonNullImmutableList<NameToken>()
				);
			}

			// Handle regular value-type constants
			var constantProperty = typeof(VBScriptConstants).GetProperty(
				builtInValueExpressionSegment.Token.Content,
				BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance
			);
			if ((constantProperty == null) || !constantProperty.CanRead || constantProperty.GetIndexParameters().Any())
				throw new NotSupportedException("Unsupported BuiltInValueToken content: " + builtInValueExpressionSegment.Token.Content);
			return new TranslatedStatementContentDetailsWithContentType(
				string.Format(
					"{0}.Constants.{1}",
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
			if (scopeAccessInformation.ParentReturnValueNameIfAny != null)
			{
				// If the segment's first (or only) member accessor and no arguments and wasn't expressed in the source code as a function call (ie. it didn't
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
				}
			}

            // TODO: Add handling for BuiltInValueToken (ie. "Err" for "Err.Raise" calls - or "Err.Description" accesses?)
            var targetBuiltInFunction = firstMemberAccessToken as BuiltInFunctionToken;
            if (targetBuiltInFunction != null)
            {
                var nameOfSupportFunction = GetNameForBuiltInFunction(targetBuiltInFunction);
                var rewrittenMemberAccessTokens = new[] { new DoNotRenameNameToken(nameOfSupportFunction, targetBuiltInFunction.LineIndex) }
                    .Concat(callExpressionSegment.MemberAccessTokens.Skip(1));
                return TranslateCallExpressionSegment(
                    _supportRefName.Name,
                    rewrittenMemberAccessTokens,
                    callExpressionSegment.Arguments,
                    scopeAccessInformation,
                    indexInCallSet: 0, // Since this is a single CallExpressionSegment the indexInCallSet value to pass is always zero
                    targetIsKnownToBeBuiltInFunction: true
                );
            }

            var targetReference = _nameRewriter.GetMemberAccessTokenName(firstMemberAccessToken);
            var result = TranslateCallExpressionSegment(
                targetReference,
                callExpressionSegment.MemberAccessTokens.Skip(1),
                callExpressionSegment.Arguments,
                scopeAccessInformation,
                indexInCallSet: 0, // Since this is a single CallExpressionSegment the indexInCallSet value to pass is always zero
                targetIsKnownToBeBuiltInFunction: false
            );
            var targetNameToken = firstMemberAccessToken as NameToken;
            if (targetNameToken != null)
            {
                result = new TranslatedStatementContentDetailsWithContentType(
                    result.TranslatedContent,
                    result.ContentType,
                    result.VariablesAccessed.Add(targetNameToken)
                );
            }
            return result;
        }

        private string GetNameForBuiltInFunction(BuiltInFunctionToken builtInFunctionToken)
        {
            if (builtInFunctionToken == null)
                throw new ArgumentNullException("builtInFunctionToken");

            var supportFunction = typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).GetMethod(
                builtInFunctionToken.Content,
                BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance
            );
            if (supportFunction == null)
                throw new NotSupportedException("Unsupported BuiltInFunctionToken content: " + builtInFunctionToken.Content);
            return supportFunction.Name;
        }

        private TranslatedStatementContentDetailsWithContentType TranslateCallExpressionSegment(
            string targetName,
            IEnumerable<IToken> targetMemberAccessTokens,
            IEnumerable<Expression> arguments,
            ScopeAccessInformation scopeAccessInformation,
            int indexInCallSet,
            bool targetIsKnownToBeBuiltInFunction)
        {
            if (string.IsNullOrWhiteSpace(targetName))
                throw new ArgumentException("Null/blank targetName specified");
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

            // If this is part of a CallSetExpression and is not the first item then there is no point trying to analyse the origin of the targetName (check
            // its scope, etc..) since this should be something of the form "_.CALL(_outer, "F", _.ARGS.Val(0))" - there is nothing to be gained from trying
            // to guess whether it's a function or what variables were accessed since this has already been done. (It's still important to check for
            // undeclared variables referenced in the arguments but that is all handled later on). The same applies if the target is known to be
            // a built-in function (such as CDate).
            DeclaredReferenceDetails targetReferenceDetailsIfAvailable;
            CSharpName nameOfTargetContainerIfRequired;
            if (targetIsKnownToBeBuiltInFunction || (indexInCallSet > 0))
            {
                targetReferenceDetailsIfAvailable = null;
                nameOfTargetContainerIfRequired = null;
            }
            else
            {
                targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(targetName, _nameRewriter);
                nameOfTargetContainerIfRequired = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(targetName, _envRefName, _outerRefName, _nameRewriter);
            }

            // If there are no member access tokens then we have to consider whether this is a function call or property access, we can find this
            // out by looking into the scope access information. Note: Further down, we rely on function / property calls being identified at this
            // point for cases where there are no target member accessors (it means that if we get further down and there are no target member
            // accessors that it must not be a function or property call).
            // - The call semantics are different for a function call, if there is a method "F" in the outermost scope then something like
            //   "_.CALL(_outer, "F", args)" would be generated but if "F" isn't a function then "_.CALL(_outer.F, new string[], args)"
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
                            "{0}.CALL({1}, \"{2}\"",
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
            // constitued the entire statement then it would have to be forced through a .VAL call but that will also have been handled
            // before this point in the TryToGetShortCutStatementResponse call in the public Translate method (see notes in the
            // TryToGetShortCutStatementResponse method for more information about this).
            if (!targetMemberAccessTokensArray.Any() && !argumentsArray.Any())
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
                "{0}.CALL({1}{2}",
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
                        _nameRewriter.GetMemberAccessTokenName(targetMemberAccessTokensArray[index]).ToLiteral()
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

            bool isConfirmedToBeByVal;
            if (forceAllArgumentsToBeByVal)
            {
                // If the arguments are for a function that is known to only take ByVal arguments (all of the built-in functions, such as CDate,
                // for example, then we can skip the hard work and set isConfirmedToBeByVal to true straight away)
                isConfirmedToBeByVal = true;
            }
            else if (argumentValue.Segments.Count() > 1)
            {
                // If there are multiple segments here then it must be ByVal (it's fairly difficult to actually not be ByVal - it basically boils
                // down to being a simple variable reference - eg. "a" - or one or more array accesses - eg. "a(0)" or "a(0, 1)" or "a(0)(1)"
                // though differentiating between the case where "a(0)" is an array accessor and where it's a default function/property
                // access is where some of the difficulty comes in). All of the below will try to decide what to do.
                isConfirmedToBeByVal = true;
            }
            else
            {
                // If there is a single segment that is a constant then it must be ByVal, same if it's a bracketed segment (VBScript treats values
                // passed wrapped in extra brackets as being ByVal), same if it's a NewInstanceExpressionSegment (there is no point allowing the
                // reference to be changed since nothing has a reference to the new instance being passed in)
                var singleSegment = argumentValue.Segments.First();
                if ((singleSegment is NumericValueExpressionSegment)
                || (singleSegment is StringValueExpressionSegment)
                || (singleSegment is BuiltInValueExpressionSegment)
                || (singleSegment is NewInstanceExpressionSegment))
                    isConfirmedToBeByVal = true;
                else if (singleSegment is BracketedExpressionSegment)
                {
                    // If this argument content is bracketed then it must be passed ByVal (a VBScript peculiarity) but if the brackets are present
                    // for that purpose only then they needn't be included in the final output, so if there is a single term that has been wrapped
                    // in brackets then unwrap the term (overwrite the argumentValue reference so that this change is reflected in the rendering
                    // call further down).
                    while ((singleSegment is BracketedExpressionSegment) && (((BracketedExpressionSegment)singleSegment).Segments.Count() == 1))
                        singleSegment = ((BracketedExpressionSegment)singleSegment).Segments.First();
                    argumentValue = new Expression(new[] { singleSegment });
                    isConfirmedToBeByVal = true;
                }
                else
                {
                    // If this is either a single CallExpressionSegment or a CallSetExpressionSegment then there are some conditions that will mean that
                    // it should definitely be passed ByVal. If there are any nested member accessors (eg. "a.Name" or "a(0).Name") then pass ByVal. If
                    // the first call segment is to a built-in function then pass ByVal
                    CallSetItemExpressionSegment initialCallSetItemExpressionSegmentToCheckIfAny;
                    var callExpressionSegment = singleSegment as CallExpressionSegment;
                    if (callExpressionSegment != null)
                        initialCallSetItemExpressionSegmentToCheckIfAny = callExpressionSegment;
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
                                initialCallSetItemExpressionSegmentToCheckIfAny = callSetExpressionSegment.CallExpressionSegments.First();
                        }
                        else
                            initialCallSetItemExpressionSegmentToCheckIfAny = null;
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
                            var rewrittenName = _nameRewriter.GetMemberAccessTokenName(initialCallSetItemExpressionSegmentToCheckIfAny.MemberAccessTokens.First());
                            var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(rewrittenName, _nameRewriter);
                            if (targetReferenceDetailsIfAvailable == null)
                            {
                                // If this is an undeclared reference then it will be implicitly declared later as a variable and so will be elligible
                                // to be passed ByRef
                                isConfirmedToBeByVal = false;
                            }
                            else
                            {
                                isConfirmedToBeByVal = (
                                    (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Function) ||
                                    (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Property)
                                );
                            }
                        }
                    }
                    else
                        isConfirmedToBeByVal = false;
                }
            }
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
            // above). On the other hand, there shouldn't be any other time of expression segment that could get this far either! (Access of a variable
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
                if (possibleByRefCallExpressionSegment.MemberAccessTokens.Count() > 2)
                    throw new NotSupportedException("Unexpected argumentValue content - didn't expect a CallExpressionSegment with multiple MemberAccessTokens at this point");
                if (!possibleByRefCallExpressionSegment.MemberAccessTokens.Any())
                    throw new NotSupportedException("Unexpected argumentValue content - didn't expect a CallExpressionSegment without any arguments at this point");
                possibleByRefTarget = Translate(
                    new CallExpressionSegment(
                        possibleByRefCallExpressionSegment.MemberAccessTokens,
                        new Expression[0],
                        CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Present
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

					possibleByRefTarget = Translate(
						new CallExpressionSegment(
							possibleByRefCallSetExpressionSegment.CallExpressionSegments.First().MemberAccessTokens,
							new Expression[0],
							CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Present
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
                        content,
                        callSetItemExpression.MemberAccessTokens,
                        callSetItemExpression.Arguments,
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

        private TranslatedStatementContentDetailsWithContentType Translate(NewInstanceExpressionSegment newInstanceExpressionSegment, LegacyParser.ScopeLocationOptions scopeLocation)
        {
            if (newInstanceExpressionSegment == null)
                throw new ArgumentNullException("newInstanceExpressionSegment");
            if (!Enum.IsDefined(typeof(LegacyParser.ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");

            return new TranslatedStatementContentDetailsWithContentType(
                string.Format(
                    "new {0}({1}, {2}, {3})",
                    _nameRewriter.GetMemberAccessTokenName(newInstanceExpressionSegment.ClassName),
                    _supportRefName.Name,
                    _envRefName.Name,
                    _outerRefName.Name
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

		private string ApplyReturnTypeGuarantee(string translatedContent, ExpressionReturnTypeOptions contentType, ExpressionReturnTypeOptions requiredReturnType, int lineIndex)
		{
			if (string.IsNullOrWhiteSpace(translatedContent))
				throw new ArgumentException("Null/blank translatedContent specified");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

			switch(requiredReturnType)
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
					if (contentType == ExpressionReturnTypeOptions.Value)
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
            var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(rewrittenName, _nameRewriter);
            if ((targetReferenceDetailsIfAvailable == null) || (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.ExternalDependency))
                rewrittenName = _envRefName.Name + "." + rewrittenName;
            else if (targetReferenceDetailsIfAvailable != null)
            {
                if ((targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Function)
                || (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Property))
                    return null;

                if (targetReferenceDetailsIfAvailable.ScopeLocation == LegacyParser.ScopeLocationOptions.OutermostScope)
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
    }
}
