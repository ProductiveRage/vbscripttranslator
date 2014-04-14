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
                    result.VariablesAccesed
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
                    result.VariablesAccesed
				);
            }

            var resultLeft = TranslateNonOperatorSegment(segments[0], scopeAccessInformation);
            var resultRight = TranslateNonOperatorSegment(segments[2], scopeAccessInformation);
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
                resultLeft.VariablesAccesed.AddRange(resultRight.VariablesAccesed)
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
                return Translate(bracketedExpressionSegment);

            var newInstanceExpressionSegment = segment as NewInstanceExpressionSegment;
            if (newInstanceExpressionSegment != null)
                return Translate(newInstanceExpressionSegment, scopeAccessInformation.ScopeLocation);

            throw new NotSupportedException("Unsupported segment type: " + segment.GetType());
        }

        private TranslatedStatementContentDetailsWithContentType Translate(BracketedExpressionSegment bracketedExpressionSegment)
		{
			if (bracketedExpressionSegment == null)
				throw new ArgumentNullException("bracketedExpressionSegment");

			throw new NotImplementedException(); // TODO
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
			// need to be replaced with the scopeAccessInformation's parentReturnValueNameIfAny value (if there is one). Note: If ParentReturnValueNameIfAny is
			// non-null then ScopeDefiningParentIfAny will also be non-null (according to the ScopeAccessInformation class).
			if (scopeAccessInformation.ParentReturnValueNameIfAny != null)
			{
				// If the segment's first (or only) member accessor and no arguments and wasn't expressed in the source code as a function call (ie. it didn't
				// have brackets after the member accessor) and the single member accessor matches the name of the ScopeDefiningParentIfAny then we need to
				// make the replacement..
				// - If arguments are specified or brackets used with no arguments then it is always a function (or property) call and the return-value
				//   replacement does not need to be made (the return value may not be DIM'd or REDIM'd to an array and so element access is not allowed,
				//   so ANY argument use always points to a function / property call)
				var rewrittenFirstMemberAccessor = _nameRewriter.GetMemberAccessTokenName(callExpressionSegment.MemberAccessTokens.First());
				var rewrittenScopeDefiningParentName = _nameRewriter.GetMemberAccessTokenName(scopeAccessInformation.ScopeDefiningParentIfAny.Name);
				if ((rewrittenFirstMemberAccessor == rewrittenScopeDefiningParentName)
				&& !callExpressionSegment.Arguments.Any()
				&& (callExpressionSegment.ZeroArgumentBracketsPresence == CallExpressionSegment.ArgumentBracketPresenceOptions.Absent))
				{
					// The ScopeDefiningParentIfAny's Name will have come from a TempValueNameGenerator rather than VBScript source code, and as such it
					// should not be passed through any VBScriptNameRewriter processing. Using a DoNotRenameNameToken means that, if the extension method
					// GetMemberAccessTokenName is consistently used for VBScriptNameRewriter access, its name won't be altered.
                    var parentReturnValueNameToken = new DoNotRenameNameToken(
                        scopeAccessInformation.ParentReturnValueNameIfAny.Name,
                        callExpressionSegment.MemberAccessTokens.First().LineIndex
                    );
					callExpressionSegment = new CallExpressionSegment(
						new[] { parentReturnValueNameToken }.Concat(callExpressionSegment.MemberAccessTokens.Skip(1)),
						new Expression[0],
						CallExpressionSegment.ArgumentBracketPresenceOptions.Absent
					);
				}
			}

            var result = TranslateCallExpressionSegment(
                _nameRewriter.GetMemberAccessTokenName(callExpressionSegment.MemberAccessTokens.First()),
                callExpressionSegment.MemberAccessTokens.Skip(1),
                callExpressionSegment.Arguments,
                scopeAccessInformation,
                0 // Since this is a single CallExpressionSegment the indexInCallSet value to pass is always zero
            );
            var targetNameToken = callExpressionSegment.MemberAccessTokens.First() as NameToken;
            if (targetNameToken != null)
            {
                result = new TranslatedStatementContentDetailsWithContentType(
                    result.TranslatedContent,
                    result.ContentType,
                    result.VariablesAccesed.Add(targetNameToken)
                );
            }
            return result;
        }

        private TranslatedStatementContentDetailsWithContentType TranslateCallExpressionSegment(
            string targetName,
            IEnumerable<IToken> targetMemberAccessTokens,
            IEnumerable<Expression> arguments,
            ScopeAccessInformation scopeAccessInformation,
            int indexInCallSet)
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
            // its scope, etc..) since this should be something of the form "_.CALL(_outer, "F", _.ARGS.Val(0).GetArgs())" - there is nothing to be gained
            // from trying to guess whether it's a function or what variables were accessed since this has already been done. (It's still important to
            // check for undeclared variables referenced in the arguments but that is all handled later on).
            ScopeAccessInformation_Extensions.DeclaredReferenceDetails targetReferenceDetailsIfAvailable;
            string nameOfTargetContainerIfRequired;
            if (indexInCallSet > 0)
            {
                targetReferenceDetailsIfAvailable = null;
                nameOfTargetContainerIfRequired = null;
            }
            else
            {
                targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(targetName, _nameRewriter);
                if (targetReferenceDetailsIfAvailable == null)
                {
                    if (scopeAccessInformation.ScopeLocation == LegacyParser.ScopeLocationOptions.WithinFunctionOrProperty)
                    {
                        // If an undeclared variable is accessed within a function (or property) then it is treated as if it was declared to be restricted
                        // to the current scope, so the nameOfTargetContainerIfRequired should be null in this case (this means that the UndeclaredVariables
                        // data returned from this process should be translated into locally-scoped DIM statements at the top of the function / property).
                        nameOfTargetContainerIfRequired = null;
                    }
                    else
                        nameOfTargetContainerIfRequired = _envRefName.Name;
                }
                else if (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.ExternalDependency)
                    nameOfTargetContainerIfRequired = _envRefName.Name;
                else if (targetReferenceDetailsIfAvailable.ScopeLocation == LegacyParser.ScopeLocationOptions.OutermostScope)
                {
                    // 2014-01-06 DWR: Used to only apply this logic if the target reference was in the OutermostScope and we were currently inside a
                    // class but I'm restructuring the outer scope so that declared variables and functions are inside a class that the outermost scope
                    // references in an identical manner to the class functions (and properties) so the outerRefName should used every time that an
                    // OutermostScope reference is accessed
                    nameOfTargetContainerIfRequired = _outerRefName.Name;
                }
                else
                    nameOfTargetContainerIfRequired = null;
            }

            // If there are no member access tokens then we have to consider whether this is a function call or property access, we can find this
            // out by looking into the scope access information. Note: Further down, we rely on function / property calls being identified at this
            // point for cases where there are no target member accessors (it means that if we get further down and there are no target member
            // accessors that it must not be a function or property call).
            // - The call semantics are different for a function call, if there is a method "F" in the outer most scope then something like
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
                            nameOfTargetContainerIfRequired ?? "this",
                            targetName
                        );
                        if (argumentsArray.Any())
                        {
                            memberCallContent.Append(", ");
                            memberCallContent.Append(_supportRefName.Name);
                            memberCallContent.Append(".ARGS");
                            for (var index = 0; index < argumentsArray.Length; index++)
                            {
                                var argumentContent = TranslateAsArgumentContent(argumentsArray[index], scopeAccessInformation);
                                memberCallContent.Append(argumentContent.TranslatedContent);
                                memberCallVariablesAccessed = memberCallVariablesAccessed.AddRange(
                                    argumentContent.VariablesAccesed
                                );
                            }
                            memberCallContent.Append(".GetArgs()");
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
                        (nameOfTargetContainerIfRequired == null) ? "" : string.Format("{0}.", nameOfTargetContainerIfRequired),
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
                (nameOfTargetContainerIfRequired == null) ? "" : string.Format("{0}.", nameOfTargetContainerIfRequired),
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
                callExpressionContent.Append(_supportRefName.Name);
                callExpressionContent.Append(".ARGS");
                for (var index = 0; index < argumentsArray.Length; index++)
                {
                    var argumentContent = TranslateAsArgumentContent(argumentsArray[index], scopeAccessInformation);
                    callExpressionContent.Append(argumentContent.TranslatedContent);
                    callExpressionVariablesAccessed = callExpressionVariablesAccessed.AddRange(
                        argumentContent.VariablesAccesed
                    );
                }
                callExpressionContent.Append(".GetArgs()");
            }

            callExpressionContent.Append(")");
			return new TranslatedStatementContentDetailsWithContentType(
				callExpressionContent.ToString(),
				ExpressionReturnTypeOptions.NotSpecified, // This could be anything so we have to report NotSpecified as the return type
                callExpressionVariablesAccessed
			);
        }

        /// <summary>
        /// This generates the calls to IBuildCallArgumentProviders (such as .Val or .Ref) based upon argument content (eg. an argument value that
        /// is the result of calling another function can never be passed ByRef whereas an argument that is a simple variable reference always
        /// CAN be passed ByRef). Note that this does not have to consider whether the target function describes arguments as ByVal or ByRef,
        /// this is just about setting up the IBuildCallArgumentProviders data so that any ByRef arguments CAN be updated on the caller
        /// where required.
        /// </summary>
        private TranslatedStatementContentDetails TranslateAsArgumentContent(Expression argumentValue, ScopeAccessInformation scopeAccessInformation)
        {
            if (argumentValue == null)
                throw new ArgumentNullException("argumentValue");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            bool isConfirmedToBeByVal;
            if (argumentValue.Segments.Count() == 1)
            {
                // If there is a single segment that is a constant then it must be ByVal
                var singleSegment = argumentValue.Segments.First();
                if ((singleSegment is NumericValueExpressionSegment)
                || (singleSegment is StringValueExpressionSegment)
                || (singleSegment is BuiltInValueExpressionSegment))
                    isConfirmedToBeByVal = true;
                else
                {
                    // If the segment is a CallExpressionSegment that refers to a built-in function or if there are member accessors (indicating
                    // property access - eg. "a.Name") then the reference can not be altered and so must be ByVal
                    var callExpressionSegment = singleSegment as CallExpressionSegment;
                    if (callExpressionSegment != null)
                    {
                        if ((callExpressionSegment.MemberAccessTokens.Count() > 1)
                        || (callExpressionSegment.MemberAccessTokens.First() is BuiltInFunctionToken))
                            isConfirmedToBeByVal = true;
                        else
                        {
                            // If the CallExpressionSegment's target is confirmed to be a function then the refrence can not be altered and so
                            // the argument must be ByVal
                            var rewrittenName = _nameRewriter.GetMemberAccessTokenName(callExpressionSegment.MemberAccessTokens.First());
                            var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(rewrittenName, _nameRewriter);
                            if (targetReferenceDetailsIfAvailable != null)
                            {
                                isConfirmedToBeByVal =
                                    (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Function) ||
                                    (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.Property);
                            }
                            else
                                isConfirmedToBeByVal = false;
                        }
                    }
                    else
                        isConfirmedToBeByVal = false;
                }
            }
            else
                isConfirmedToBeByVal = false;
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
                    translatedCallExpressionByValArgumentContent.VariablesAccesed
                );
            }

            bool isConfirmedToBeByRef;
            if (argumentValue.Segments.Count() == 1)
            {
                // If the CallExpressionSegment's target is confirmed to not be a function and not have any member accessors or arguments then
                // it CAN be passed ByRef
                var callExpressionSegment = argumentValue.Segments.First() as CallExpressionSegment;
                if ((callExpressionSegment != null) && (callExpressionSegment.MemberAccessTokens.Count() == 1) && !callExpressionSegment.Arguments.Any())
                {
                    var rewrittenName = _nameRewriter.GetMemberAccessTokenName(callExpressionSegment.MemberAccessTokens.First());
                    var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(rewrittenName, _nameRewriter);
                    if (targetReferenceDetailsIfAvailable == null)
                    {
                        // If this is an undeclared reference then it will be implicitly declared later as a variable and so WILL be elligible
                        // to be passed ByRef (it can't be a function or property which is what would get it off the hook)
                        isConfirmedToBeByRef = true;
                    }
                    else
                    {
                        isConfirmedToBeByRef =
                            (targetReferenceDetailsIfAvailable.ReferenceType != ReferenceTypeOptions.Function) &&
                            (targetReferenceDetailsIfAvailable.ReferenceType != ReferenceTypeOptions.Property);
                    }
                }
                else
                    isConfirmedToBeByRef = false;
            }
            else
                isConfirmedToBeByRef = false;
            if (isConfirmedToBeByRef)
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
                    translatedCallExpressionByRefArgumentContent.VariablesAccesed
                );
            }

            throw new NotImplementedException(); // TODO: Deal with maybe-ByRef

            // Single CallExpressionSegment where single target is confirmed to NOT be a function and there are no arguments => ByRef *********

            // TODO: Fix support for "a(0)(1)" (identified as CallExpressionSegment, NumericValueSegment.. maybe this IS correct?!)
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
                        index
                    );
                }
                
                // Overwrite any previous string content since it effectively gets passed through as an accumulator
                content = translatedContent.TranslatedContent;
                variablesAccessed = variablesAccessed.AddRange(translatedContent.VariablesAccesed);
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
