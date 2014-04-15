using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using StageTwoParser = VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation.StatementTranslation
{
	public class ValueSettingsStatementsTranslator : ITranslateValueSettingsStatements
    {
		private readonly CSharpName _supportRefName, _envRefName, _outerRefName;
		private readonly VBScriptNameRewriter _nameRewriter;
		private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
		public ValueSettingsStatementsTranslator(
            CSharpName supportRefName,
            CSharpName envRefName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            ITranslateIndividualStatements statementTranslator,
            ILogInformation logger)
		{
			if (supportRefName == null)
				throw new ArgumentNullException("supportRefName");
            if (envRefName == null)
                throw new ArgumentNullException("envRefName");
            if (outerRefName == null)
                throw new ArgumentNullException("outerRefName");
			if (nameRewriter == null)
				throw new ArgumentNullException("nameRewriter");
			if (statementTranslator == null)
				throw new ArgumentNullException("statementTranslator");
            if (logger == null)
                throw new ArgumentNullException("logger");

			_supportRefName = supportRefName;
            _envRefName = envRefName;
            _outerRefName = outerRefName;
			_nameRewriter = nameRewriter;
			_statementTranslator = statementTranslator;
            _logger = logger;
		}

		/// <summary>
		/// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
		/// </summary>
		public TranslatedStatementContentDetails Translate(ValueSettingStatement valueSettingStatement, ScopeAccessInformation scopeAccessInformation)
		{
			if (valueSettingStatement == null)
				throw new ArgumentNullException("valueSettingStatement");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var assignmentFormatDetails = GetAssignmentFormatDetails(valueSettingStatement, scopeAccessInformation);

			var translatedExpressionContentDetails = _statementTranslator.Translate(
				valueSettingStatement.Expression,
				scopeAccessInformation,
				(valueSettingStatement.ValueSetType == ValueSettingStatement.ValueSetTypeOptions.Set)
					? ExpressionReturnTypeOptions.Reference
					: ExpressionReturnTypeOptions.Value
			);

			return new TranslatedStatementContentDetails(
				assignmentFormatDetails.AssigmentFormat(translatedExpressionContentDetails.TranslatedContent),
				assignmentFormatDetails.VariablesAccessed.AddRange(translatedExpressionContentDetails.VariablesAccessed)
			);
		}

		private ValueSettingStatementAssigmentFormatDetails GetAssignmentFormatDetails(ValueSettingStatement valueSettingStatement, ScopeAccessInformation scopeAccessInformation)
		{
			if (valueSettingStatement == null)
				throw new ArgumentNullException("valueSettingStatement");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			// The ValueToSet content should be reducable to a single expression segment; a CallExpression or a CallSetExpression
			var targetExpression = ExpressionGenerator.Generate(valueSettingStatement.ValueToSet.BracketStandardisedTokens).ToArray();
			if (targetExpression.Length != 1)
				throw new ArgumentException("The ValueToSet should always be described by a single expression");
			var targetExpressionSegments = targetExpression[0].Segments.ToArray();
			if (targetExpressionSegments.Length != 1)
				throw new ArgumentException("The ValueToSet should always be described by a single expression containing a single segment");

			// If there is only a single CallExpressionSegment with a single token then that token must be a NameToken otherwise the statement would be
			// trying to assign a value to a constant or keyword or something inappropriate. If it IS a NameToken, then this is the easiest case - it's
			// a simple assignment, no SET method call required.
			var expressionSegment = targetExpressionSegments[0];
			var callExpressionSegment = expressionSegment as CallExpressionSegment;
			if ((callExpressionSegment != null) && (callExpressionSegment.MemberAccessTokens.Take(2).Count() < 2) && !callExpressionSegment.Arguments.Any())
			{
				var singleTokenAsName = callExpressionSegment.MemberAccessTokens.Single() as NameToken;
				if (singleTokenAsName == null)
					throw new ArgumentException("Where a ValueSettingStatement's ValueToSet expression is a single expression with a single CallExpressionSegment with one token, that token must be a NameToken");

				// If this single token is the function name (if we're in a function or property) then we need to make the ParentReturnValueNameIfAny
				// replacement so that the return value reference is updated.
				// TODO: If callExpressionSegment.ZeroArgumentBracketsPresence == Present then throw a "Type mismatch" exception rather than make
				// the replacement (in order to be consistent with VBScript's runtime behaviour)
				var rewrittenFirstMemberAccessor = _nameRewriter.GetMemberAccessTokenName(singleTokenAsName);
				var isSingleTokenSettingParentScopeReturnValue = (
					(scopeAccessInformation.ParentReturnValueNameIfAny != null) &&
					rewrittenFirstMemberAccessor == _nameRewriter.GetMemberAccessTokenName(scopeAccessInformation.ScopeDefiningParentIfAny.Name)
				);

                // If the "targetAccessor" is an undeclared variable then it must be accessed through the envRefName (this is a reference that should
                // be passed into the containing class' constructor since C# doesn't support the concept of abritrary unintialised references)
                var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(rewrittenFirstMemberAccessor, _nameRewriter);
                if ((targetReferenceDetailsIfAvailable == null) || (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.ExternalDependency))
                    rewrittenFirstMemberAccessor = _envRefName.Name + "." + rewrittenFirstMemberAccessor;
                else if (targetReferenceDetailsIfAvailable.ScopeLocation == ScopeLocationOptions.OutermostScope)
                    rewrittenFirstMemberAccessor = _outerRefName.Name + "." + rewrittenFirstMemberAccessor;
                return new ValueSettingStatementAssigmentFormatDetails(
					translatedExpression => string.Format(
						"{0} = {1}",
						isSingleTokenSettingParentScopeReturnValue
							? scopeAccessInformation.ParentReturnValueNameIfAny.Name
							: rewrittenFirstMemberAccessor,
						translatedExpression
					),
					new NonNullImmutableList<NameToken>(new[] { singleTokenAsName })
				);
			}

			// If this is a more complicated case then the single assignment case covered above then we need to break it down into a target reference,
			// an optional single member accessor and zero or more arguments. For example -
			//
			//  "a(0)"                 ->  "a", null, [0]
			//  "a.Role(0)"            ->  "a", "Role", [0]
			//  "a.b.Role(0)"          ->  "a.b", "Role", [0]
			//  "a(0).Name"            ->  "a(0)", "Name", []
			//  "c.r.Fields(0).Value"  ->  "c.r.Fields(0)", "Value", []
			//
			// This allows the StatementTranslator to handle reference access until the last moment, providing the logic to access the "target". Then
			// there is much less to do when determining how to set the value on the target. The case where there are arguments is the only time that
			// defaults need to be considered (eg. "a(0)" could be an array access if "a" is an array or it could be a default property access if "a"
			// has a default indexed property).

			// Get a set of CallExpressionSegments..
			List<CallSetItemExpressionSegment> callExpressionSegments;
			if (callExpressionSegment != null)
                callExpressionSegments = new List<CallSetItemExpressionSegment> { callExpressionSegment };
			else
			{
				var callSetExpressionSegment = expressionSegment as CallSetExpressionSegment;
				if (callSetExpressionSegment == null)
					throw new ArgumentException("The ValueToSet should always be described by a single expression containing a single segment, of type CallExpressionSegment or CallSetExpressionSegment");
				callExpressionSegments = callSetExpressionSegment.CallExpressionSegments.ToList();
			}

			// If the last CallExpressionSegment has more than two member accessor then it needs breaking up since we need a target, at most a single
			// member accessor against that target and arguments. So if there are more than two member accessors in the last segments then we want to
			// split it so that all but one are in one segment (with no arguments) and then the last one (to go WITH the arguments) in the last entry.
			var numberOfMemberAccessTokensInLastCallExpressionSegment = callExpressionSegments.Last().MemberAccessTokens.Count();
			if (numberOfMemberAccessTokensInLastCallExpressionSegment > 2)
			{
				var lastCallExpressionSegments = callExpressionSegments.Last();
				callExpressionSegments.RemoveAt(callExpressionSegments.Count - 1);
				callExpressionSegments.Add(
					new CallExpressionSegment(
						lastCallExpressionSegments.MemberAccessTokens.Take(numberOfMemberAccessTokensInLastCallExpressionSegment - 1),
						new StageTwoParser.Expression[0],
						CallExpressionSegment.ArgumentBracketPresenceOptions.Absent // Can't be any brackets as we're splitting a CallExpressionSegment in two
					)
				);
				callExpressionSegments.Add(
					new CallExpressionSegment(
						new[] { lastCallExpressionSegments.MemberAccessTokens.Last() },
						lastCallExpressionSegments.Arguments,
						lastCallExpressionSegments.ZeroArgumentBracketsPresence
					)
				);
			}

			// Now we either have a single segment in the set that has no more than two member accessors (lending itself to easy extraction of a
			// target and optional member accessor) or we have multiple segments where all but the last one define the target and the last entry
			// has the optional member accessor and any arguments.
			string targetAccessor;
			string optionalMemberAccessor;
			IEnumerable<StageTwoParser.Expression> arguments;
			if (callExpressionSegments.Count == 1)
			{
				// The single CallExpressionSegment may have one or two member accessors
				targetAccessor = _nameRewriter.GetMemberAccessTokenName(callExpressionSegments[0].MemberAccessTokens.First());
				if (callExpressionSegments[0].MemberAccessTokens.Count() == 1)
					optionalMemberAccessor = null;
				else
				{
					optionalMemberAccessor = _nameRewriter.GetMemberAccessTokenName(
						callExpressionSegments[0].MemberAccessTokens.Skip(1).Single()
					);
				}
				arguments = callExpressionSegments[0].Arguments;

                var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(targetAccessor, _nameRewriter);
                if (targetReferenceDetailsIfAvailable == null)
                {
                    // If an undeclared variable is accessed within a function (or property) then it is treated as if it was declared to be restricted
                    // to the current scope, so the targetAccessor does not require a prefix in this case (this means that the UndeclaredVariables data
                    // returned from this process should be translated into locally-scoped DIM statements at the top of the function / property).
                    if (scopeAccessInformation.ScopeLocation != ScopeLocationOptions.WithinFunctionOrProperty)
                        targetAccessor = _envRefName.Name + "." + targetAccessor;
                }
                else if (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.ExternalDependency)
                    targetAccessor = _envRefName.Name + "." + targetAccessor;
                else if (targetReferenceDetailsIfAvailable.ScopeLocation == ScopeLocationOptions.OutermostScope)
                    targetAccessor = _outerRefName.Name + "." + targetAccessor;
                else
                {
                    // If this single token is the function name (if we're in a function or property) then we need to make the ParentReturnValueNameIfAny
                    // replacement so that the return value reference is updated.
                    // TODO: If callExpressionSegment.ZeroArgumentBracketsPresence == Present then throw a "Type mismatch" exception rather than make
                    // the replacement (in order to be consistent with VBScript's runtime behaviour)
                    var isSingleTokenSettingParentScopeReturnValue = (
                        (scopeAccessInformation.ParentReturnValueNameIfAny != null) &&
                        targetAccessor == _nameRewriter.GetMemberAccessTokenName(scopeAccessInformation.ScopeDefiningParentIfAny.Name)
                    );
                    if (isSingleTokenSettingParentScopeReturnValue)
                        targetAccessor = scopeAccessInformation.ParentReturnValueNameIfAny.Name;
                }
			}
			else
			{
				var targetAccessCallExpressionSegments = callExpressionSegments.Take(callExpressionSegments.Count() - 1);
				var targetAccessExpressionSegments = (targetAccessCallExpressionSegments.Count() > 1)
					? new IExpressionSegment[] { new CallSetExpressionSegment(targetAccessCallExpressionSegments) }
					: new IExpressionSegment[] { targetAccessCallExpressionSegments.Single() };
				targetAccessor =
					_statementTranslator.Translate(
						new StageTwoParser.Expression(targetAccessExpressionSegments),
						scopeAccessInformation,
						ExpressionReturnTypeOptions.NotSpecified
					).TranslatedContent;

				// The last CallExpressionSegment may only have one member accessor
				var lastCallExpressionSegment = callExpressionSegments.Last();
				optionalMemberAccessor = _nameRewriter.GetMemberAccessTokenName(lastCallExpressionSegment.MemberAccessTokens.Single());
				arguments = lastCallExpressionSegment.Arguments;

				// Note: In this case, we don't have to apply any special logic to make "return value replacements" for assignment targets
				// (when "F2 = 1" is setting the return value for the function "F2" that we're inside of, for example) since this will be
				// handled by the statement translator. The cases above where there was only a single token or only a single call
				// expression segment needed additional logic since they don't use the statement translator for the left hand
				// side of the assignment expression.
			}

			// Regardless of how we've gone about trying to access the data in the callExpressionSegments data above, we now need to get a
			// set of all variables that are accessed so that later on we can identify any undeclared variable access attempts
			// - TODO: If manipulated segments to include function return value, it shouldn't affect the variablesAccessed retrieval but
			//   need to make a note explaining how/why
            IEnumerable<NameToken> variablesAccessed;
            if (!callExpressionSegments.Any())
                variablesAccessed = new NameToken[0];
            else
            {
                // We will have one or more CallSetItemExpressionSegment instances, the first of which will always have at least one
                // Member Access Token since the segments were initially extracted from a CallExpressionSegment (which always has at
                // least one) or from a CallSetExpressionSegment (whose first segment always has at least one). This means that we
                // can always reform the segment(s) back into either CallExpressionSegment or CallSetExpressionSegment, which
                // can then be passed to the statementTranslator to analyse for accessed variables.
                IExpressionSegment expressionToAnalyseForVariablesAccessed;
                if (callExpressionSegments.Count() == 1)
                {
                    expressionToAnalyseForVariablesAccessed = new CallExpressionSegment(
                        callExpressionSegments.First().MemberAccessTokens,
                        callExpressionSegments.First().Arguments,
                        callExpressionSegments.First().ZeroArgumentBracketsPresence
                    );
                }
                else
                    expressionToAnalyseForVariablesAccessed = new CallSetExpressionSegment(callExpressionSegments);
                variablesAccessed = _statementTranslator.Translate(
                        new StageTwoParser.Expression(new[] { expressionToAnalyseForVariablesAccessed }),
                        scopeAccessInformation,
                        ExpressionReturnTypeOptions.NotSpecified
                    )
                    .VariablesAccessed;
            }

            // Note: The translatedExpression will already account for whether the statement is of type LET or SET
            var argumentsContent = _statementTranslator.TranslateAsArgumentProvider(arguments, scopeAccessInformation);
            variablesAccessed = variablesAccessed.Concat(argumentsContent.VariablesAccessed);
            return new ValueSettingStatementAssigmentFormatDetails(
				translatedExpression => string.Format(
					"{0}.SET({1}, {2}, {3}, {4})",
					_supportRefName.Name,
					targetAccessor,
					(optionalMemberAccessor == null) ? "null" : optionalMemberAccessor.ToLiteral(),
                    argumentsContent.TranslatedContent,
					translatedExpression
				),
				variablesAccessed.ToNonNullImmutableList()
			);
		}

		private class ValueSettingStatementAssigmentFormatDetails
		{
			public ValueSettingStatementAssigmentFormatDetails(Func<string, string> assigmentFormat, NonNullImmutableList<NameToken> variablesAccessed)
			{
				if (assigmentFormat == null)
					throw new ArgumentNullException("assigmentFormat");
				if (variablesAccessed == null)
					throw new ArgumentNullException("memberCallVariablesAccessed");

				AssigmentFormat = assigmentFormat;
				VariablesAccessed = variablesAccessed;
			}

			/// <summary>
			/// This will never be null
			/// </summary>
			public Func<string, string> AssigmentFormat { get; private set; }

			/// <summary>
			/// This will never be null
			/// </summary>
			public NonNullImmutableList<NameToken> VariablesAccessed { get; private set; }
		}
	}
}
