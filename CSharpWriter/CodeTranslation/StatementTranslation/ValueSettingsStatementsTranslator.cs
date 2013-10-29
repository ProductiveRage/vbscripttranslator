using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
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
		private readonly CSharpName _supportClassName;
		private readonly VBScriptNameRewriter _nameRewriter;
		private readonly ITranslateIndividualStatements _statementTranslator;
		public ValueSettingsStatementsTranslator(CSharpName supportClassName, VBScriptNameRewriter nameRewriter, ITranslateIndividualStatements statementTranslator)
		{
			if (supportClassName == null)
				throw new ArgumentNullException("supportClassName");
			if (nameRewriter == null)
				throw new ArgumentNullException("nameRewriter");
			if (statementTranslator == null)
				throw new ArgumentNullException("statementTranslator");

			_supportClassName = supportClassName;
			_nameRewriter = nameRewriter;
			_statementTranslator = statementTranslator;
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


			throw new NotImplementedException(); // TODO
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

				// TODO: Explain..
				var isSingleTokenSettingParentScopeReturnValue = (
					(scopeAccessInformation.ScopeDefiningParentIfAny != null) &&
					(scopeAccessInformation.ParentReturnValueNameIfAny != null) &&
					scopeAccessInformation.ScopeDefiningParentIfAny.Name.Content.Equals(singleTokenAsName.Content, StringComparison.InvariantCultureIgnoreCase)
				);
				return new ValueSettingStatementAssigmentFormatDetails(
					translatedExpression => string.Format(
						"{0} = {1}",
						isSingleTokenSettingParentScopeReturnValue
							? scopeAccessInformation.ParentReturnValueNameIfAny.Name
							: _nameRewriter(singleTokenAsName).Name,
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
			List<CallExpressionSegment> callExpressionSegments;
			if (callExpressionSegment != null)
				callExpressionSegments = new List<CallExpressionSegment> { callExpressionSegment };
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
						CallExpressionSegment.ArgumentBracketPresenceOptions.Absent // TODO: Do properly
					)
				);
				callExpressionSegments.Add(
					new CallExpressionSegment(
						new[] { lastCallExpressionSegments.MemberAccessTokens.Last() },
						lastCallExpressionSegments.Arguments,
						CallExpressionSegment.ArgumentBracketPresenceOptions.Absent // TODO: Do properly
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
				// TODO: Need to consider whether we're in a function or property and whether this is an assigment for its return value
				// - If the first member accessor's name when passed through the nameRewriter matches "scopeAccessInformation.ScopeDefiningParentIfAny
				//   .Name" (if there is a scope-defining parent) then use the "scopeAccessInformation.parentReturnValueNameIfAny" (if non-null)

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
			}
			else
			{
				// TODO: Need to consider whether we're in a function or property and whether this is an assigment for its return value
				// - If the first call expression segment's member accessor's name when passed through the nameRewriter matches
				//   "scopeAccessInformation.ScopeDefiningParentIfAny.Name" (if there is a scope-defining parent) then use the
				//   "scopeAccessInformation.parentReturnValueNameIfAny" (if non-null)
				// ******* If the returnVal logic has to go into the StatementTranslator then we don't have to do any special handling
				//         here since that will take care of it (the only thing it doesn't handle is the lastCallExpressionSegment
				//         which can't be applicable for a replacement as it has to be a property of function hanging off a
				//         reference (since we know there are more than two segments) and we can't do return value
				//         replacing in that scenario

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
			}

			// Regardless of how we've gone about trying to access the data in the callExpressionSegments data above, we now need to get a
			// set of all variables that are accessed so that later on we can identify any undeclared variable access attempts
			// - TODO: If manipulated segments to include function return value, it shouldn't affect the variablesAccessed retrieval but
			//   need to make a note explaining how/why
			var variablesAccessed = GetAccessedVariables(callExpressionSegments, scopeAccessInformation);

			// Note: The translatedExpression will already account for whether the statement is of type LET or SET
			// TODO: Explain..
			var isCallExpressionSettingParentScopeReturnValue = (
				(scopeAccessInformation.ScopeDefiningParentIfAny != null) &&
				(scopeAccessInformation.ParentReturnValueNameIfAny != null) &&
				scopeAccessInformation.ScopeDefiningParentIfAny.Name.Content.Equals(targetAccessor, StringComparison.InvariantCultureIgnoreCase)
			);
			return new ValueSettingStatementAssigmentFormatDetails(
				translatedExpression => string.Format(
					"{0}.SET({1}, {2}, {3}, {4})",
					_supportClassName.Name,
					targetAccessor,
					(optionalMemberAccessor == null) ? "null" : optionalMemberAccessor.ToLiteral(),
					arguments.Any()
						? string.Format(
							"new object[] {{ {0} }}",
							string.Join(
								", ",
								arguments.Select(a => _statementTranslator.Translate(a, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified).TranslatedContent)
							)
						)
						: "new object[0]",
					translatedExpression
				),
				variablesAccessed.ToNonNullImmutableList()
			);
		}

		private IEnumerable<NameToken> GetAccessedVariables(IEnumerable<CallExpressionSegment> callExpressionSegments, ScopeAccessInformation scopeAccessInformation)
		{
			if (callExpressionSegments == null)
				throw new ArgumentNullException("callExpressionSegments");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");

			var callExpressionSegmentsArray = callExpressionSegments.ToArray();
			if (callExpressionSegmentsArray.Any(e => e == null))
				throw new ArgumentException("Null reference encountered in callExpressionSegments set");

			if (!callExpressionSegmentsArray.Any())
				return new NameToken[0];

			var expression = (callExpressionSegmentsArray.Length == 1)
				? new StageTwoParser.Expression(callExpressionSegmentsArray)
				: new StageTwoParser.Expression(new[] { new CallSetExpressionSegment(callExpressionSegmentsArray) });
			return
				_statementTranslator.Translate(
					expression,
					scopeAccessInformation,
					ExpressionReturnTypeOptions.NotSpecified
				).VariablesAccesed;
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
