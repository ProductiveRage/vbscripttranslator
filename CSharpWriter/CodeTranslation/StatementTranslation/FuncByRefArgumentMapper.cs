using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions;
using VBScriptTranslator.CSharpWriter.Lists;
using VBScriptTranslator.CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation
{
	/// <summary>
	/// If a statement needs to pass as a by-ref argument to a function or property call a reference that is a by-ref argument of the function that the statement is within
	/// (meaning this only applies to statements that are within functions) then a temporary reference must be used since by-ref arguments are handled with lambdas and
	/// ref function arguments may not be manipulated within lambdas in C#. Similar logic is also required for code that is evaluated within an error-trapping lambda,
	/// ref function arguments must also be mapped onto aliases in order for those lambdas to be valid - in this case, however, after the work is performed, any
	/// changes to the alias reference do not need to be mapped back over the original function argument.
	/// </summary>
	public class FuncByRefArgumentMapper
	{
		private readonly VBScriptNameRewriter _nameRewriter;
		private readonly TempValueNameGenerator _tempNameGenerator;
		private readonly ILogInformation _logger;
		public FuncByRefArgumentMapper(VBScriptNameRewriter nameRewriter, TempValueNameGenerator tempNameGenerator, ILogInformation logger)
		{
			if (tempNameGenerator == null)
				throw new ArgumentNullException("tempNameGenerator");
			if (logger == null)
				throw new ArgumentNullException("logger");

			_nameRewriter = nameRewriter;
			_tempNameGenerator = tempNameGenerator;
			_logger = logger;
		}

		/// <summary>
		/// If we're within a function that has a ByRef argument a0 and that argument is passed into another function, where that argument is ByRef, then we will encounter
		/// problems since the mechanism for dealing with ByRef arguments is to generate an arguments handler with lambdas for updating the variable after the call completes,
		/// but "ref" arguments may not be included in lambdas in C#. So the a0 reference must be copied into a temporary variable that is updated after the second function
		/// call ends (even if it errors, since the argument may have been altered before the error). This function will identify which variables in an expression must be
		/// rewritten in this manner. Similar logic must be applied if a0 is referenced within error-trapped code since the translated code in that case will also be executed
		/// in a lambda - in that case, however, a0 will not need to be overwritten by the alias after the work completes, it is a "read only" mapping. If the specified
		/// scopeAccessInformation reference does not indicate that the expression is within a function (or property) then there will be work to perform.
		/// </summary>
		public NonNullImmutableList<FuncByRefMapping> GetByRefArgumentsThatNeedRewriting(
			Expression expression,
			ScopeAccessInformation scopeAccessInformation,
			NonNullImmutableList<FuncByRefMapping> rewrittenReferences)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (rewrittenReferences == null)
				throw new ArgumentNullException("rewrittenReferences");

			return GetByRefArgumentsThatNeedRewriting(expression, scopeAccessInformation, rewrittenReferences, expressionIsDirectArgumentValue: false);
		}

		private NonNullImmutableList<FuncByRefMapping> GetByRefArgumentsThatNeedRewriting(
			Expression expression,
			ScopeAccessInformation scopeAccessInformation,
			NonNullImmutableList<FuncByRefMapping> rewrittenReferences,
			bool expressionIsDirectArgumentValue)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (rewrittenReferences == null)
				throw new ArgumentNullException("rewrittenReferences");

			// If we're not within a function or property then there's no work to do since there can be no ByRef arguments in the parent construct to worry about
			var containingFunctionOrProperty = scopeAccessInformation.ScopeDefiningParent as VBScriptTranslator.LegacyParser.CodeBlocks.Basic.AbstractFunctionBlock;
			if (containingFunctionOrProperty == null)
				return rewrittenReferences;
			
			// If the containing function/property doesn't have any ByRef arguments then, again, there is nothing to worry about
			var byRefArguments = containingFunctionOrProperty.Parameters.Where(p => p.ByRef).Select(p => p.Name).ToNonNullImmutableList();
			if (!byRefArguments.Any())
				return rewrittenReferences;

			// By-ref arguments of the containing function need to be aliased when referenced within a lambda (such as when processed with HANDLEERROR) - but, depending
			// upon how they are used, this aliasing can be a one-way affair. If "a" is a by-ref function argument and the lambda-wrapped statement is "F1(a + 1)" then
			// it's not possible for F1 to change the value of "a" and so the alias value does not need to be copied back onto "a" after the lambda has been called
			// (this is a read-only by-ref argument mapping). However, if the lambda-wrapped statement is "F1(a)" then it IS possible for F1 to change the value
			// of "a" - in this case the alias value must be copied over "a" after the lamdba completes (even if it errors, in case the value of the F1 argument
			// was changed before the error occurred). The GetByRefArgumentsThatNeedRewriting below examines arguments passed in to functions to determiner whether
			// any mappings should be read-only or read-write but it will miss an important case - if the "expression" at this point is a simple reference to a
			// by-ref argument then it may be the left hand side of a value-setting statement and that will definitely need to be a read-write alias mapping
			// (otherwise the alias value would be updated within the lambda, after the value-setting was executed, but it wouldn't be mapped back on to the
			// by-ref argument, which would defeat the point of the value-setting statement).
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				var nameTokenTargetIfApplicable = TryToGetExpressionAsSingleNameToken(expression);
				if ((nameTokenTargetIfApplicable != null) && byRefArguments.Any(a => a.Content.Equals(nameTokenTargetIfApplicable.Content, StringComparison.OrdinalIgnoreCase)))
					expressionIsDirectArgumentValue = true;
			}

			return GetByRefArgumentsThatNeedRewriting(expression.Segments, byRefArguments, scopeAccessInformation, rewrittenReferences, expressionIsDirectArgumentValue);
		}

		private NonNullImmutableList<FuncByRefMapping> GetByRefArgumentsThatNeedRewriting(
			IEnumerable<IExpressionSegment> expressionSegments,
			NonNullImmutableList<NameToken> byRefArguments,
			ScopeAccessInformation scopeAccessInformation,
			NonNullImmutableList<FuncByRefMapping> rewrittenReferences,
			bool expressionIsDirectArgumentValue)
		{
			if (expressionSegments == null)
				throw new ArgumentNullException("expressionSegments");
			if (byRefArguments == null)
				throw new ArgumentNullException("byRefArguments");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (rewrittenReferences == null)
				throw new ArgumentNullException("rewrittenReferences");

			if (expressionIsDirectArgumentValue && (expressionSegments.Count() > 1))
			{
				// If an expression has multiple segments then no one part of it is elligible to be passed ByRef as a function argument - eg. "a + b" will not be
				// passed ByRef since that expression "a + b" can not be altered by the receiving function even if the argument IS marked as ByRef.
				expressionIsDirectArgumentValue = false;
			}

			var rewrittenExpressionSegments = new List<IExpressionSegment>();
			foreach (var expressionSegment in expressionSegments)
			{
				if (expressionSegment == null)
					throw new ArgumentException("Null reference encountered in expressionSegments set");

				if ((expressionSegment is BuiltInValueExpressionSegment)
				|| (expressionSegment is DateValueExpressionSegment)
				|| (expressionSegment is NewInstanceExpressionSegment)
				|| (expressionSegment is NumericValueExpressionSegment)
				|| (expressionSegment is OperationExpressionSegment)
				|| (expressionSegment is RuntimeErrorExpressionSegment)
				|| (expressionSegment is StringValueExpressionSegment))
				{
					rewrittenExpressionSegments.Add(expressionSegment);
					continue;
				}

				var bracketedExpressionSegment = expressionSegment as BracketedExpressionSegment;
				if (bracketedExpressionSegment != null)
				{
					// If an expression is wrapped in brackets then, even if it would otherwise be elligible to be passed as a direct / ByRef argument to a function,
					// the brackets mean that it won't be - eg. for the complete statement "F2 (a)", the brackets are interpreted to mean force-a-to-be-passed-ByVal
					// and so we don't want it to be considered ByRef.
					rewrittenReferences = GetByRefArgumentsThatNeedRewriting(
						bracketedExpressionSegment.Segments,
						byRefArguments,
						scopeAccessInformation,
						rewrittenReferences,
						expressionIsDirectArgumentValue: false
					);
					continue;
				}

				var callExpressionSegment = expressionSegment as CallExpressionSegment;
				if (callExpressionSegment != null)
				{
					// The CallExpressionSegment is essentially a specialised version of the CallSetItemExpressionSegment where there must be at least one
					// Member Access Token, so we can call   and then use the result to create a new CallExpressionSegment
					rewrittenReferences = RewriteCallSetItemExpressionSegment(
						callExpressionSegment,
						byRefArguments,
						scopeAccessInformation,
						rewrittenReferences,
						expressionIsDirectArgumentValue,
						callSetItemIndex: 0,
						callSetItemCount: 1
					);
					continue;
				}

				var callSetExpressionSegment = expressionSegment as CallSetExpressionSegment;
				if (callSetExpressionSegment != null)
				{
					// This is a just a combined set of CallSetItemExpressionSegment instances, so we can handle this expression segment type with more recursion
					foreach (var indexedCallExpressionSegment in callSetExpressionSegment.CallExpressionSegments.Select((s, i) => new { Segment = s, Index = i }))
					{
						rewrittenReferences = RewriteCallSetItemExpressionSegment(
							indexedCallExpressionSegment.Segment,
							byRefArguments,
							scopeAccessInformation,
							rewrittenReferences,
							expressionIsDirectArgumentValue,
							callSetItemIndex: indexedCallExpressionSegment.Index,
							callSetItemCount: callSetExpressionSegment.CallExpressionSegments.Count()
						);
					}
					continue;
				}

				throw new NotSupportedException("Unsupported expression segment type: " + expressionSegment.GetType());
			}
			return rewrittenReferences;
		}

		private NonNullImmutableList<FuncByRefMapping> RewriteCallSetItemExpressionSegment(
			CallSetItemExpressionSegment callSetItemExpressionSegment,
			NonNullImmutableList<NameToken> byRefArguments,
			ScopeAccessInformation scopeAccessInformation,
			NonNullImmutableList<FuncByRefMapping> rewrittenReferences,
			bool expressionIsDirectArgumentValue,
			int callSetItemIndex,
			int callSetItemCount)
		{
			if (callSetItemExpressionSegment == null)
				throw new ArgumentNullException("callSetItemExpressionSegment");
			if (byRefArguments == null)
				throw new ArgumentNullException("byRefArguments");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (rewrittenReferences == null)
				throw new ArgumentNullException("mutableRewrittenReferences");
			if (callSetItemIndex < 0)
				throw new ArgumentOutOfRangeException("callSetItemIndex");
			if (callSetItemCount < 1)
				throw new ArgumentOutOfRangeException("callSetItemCount");
			if (callSetItemIndex > (callSetItemCount - 1))
				throw new ArgumentOutOfRangeException("callSetItemIndex", "outside of the callSetItemCount bounds");

			// Check for cases where a ByRef argument of the containing function/property is being passed ByRef into another function/property (in this case a mapping will need to be
			// recorded that is not a read-only mapping; after the call expression is evaluated, any changes to the alias reference must be mapped back onto the source function argument)
			// - Check memberAccessTokens, but only if this is the first segment and if there is only a single token. If there are multiple segments then it would never be ByRef and so
			//   we don't need to worry about it - eg. if "a" is a ByRef argument of the containing function that is then passed into another function as a ByRef argument as "a" then
			//   it must be flagged as requiring a temporary mapping, but if it is passed as "a.Name" then it does not need to be, since the second function can not alter the value
			//   of "a". There is some spreading of logic here, that this level of knowledge is required here as well as in the statement translation process.
			// Note: The ByRef argument must be being passed DIRECTLY into another function as a ByRef argument, otherwise the ByRef rules do not apply - eg. if "a.Name" is passed
			// into a function as ByRef argument, that function can not affect the value of the expression "a.Name", nor can it affect the value of "a". There sometimes-exception
			// is expressions such as "a(0)" which may be altered when passed as a ByRef argument value, but only if "a" is an array and not if "a(0)" is a default member access.
			if (expressionIsDirectArgumentValue && (callSetItemCount == 1) && (callSetItemExpressionSegment.MemberAccessTokens.Count() == 1))
			{
				var targetAsNameToken = callSetItemExpressionSegment.MemberAccessTokens.Single() as NameToken;
				if (targetAsNameToken != null)
				{
					if (IsInByRefArgumentSet(byRefArguments, targetAsNameToken))
					{
						if (!IsAlreadyAccountedFor(rewrittenReferences, targetAsNameToken, onlyConsiderNonReadOnlyMappings: true))
						{
							// Note: If there is already a mapping for this reference but it was identified as a read-only mapping then we want to remove it since the non-read-only
							// mapping that we have identified here needs to take precedence
							if (IsAlreadyAccountedFor(rewrittenReferences, targetAsNameToken, onlyConsiderNonReadOnlyMappings: false))
								rewrittenReferences = RemoveMappingForNameToken(rewrittenReferences, targetAsNameToken);
							var rewrittenReferenceName = _tempNameGenerator(new CSharpName("byrefalias"), scopeAccessInformation);
							rewrittenReferences = rewrittenReferences.Add(new FuncByRefMapping(targetAsNameToken, rewrittenReferenceName, mappedValueIsReadOnly: false));
						}
					}
				}
			}
			
			// Now, if error-trapping may be enabled, check for cases where a ByRef argument of the containing function/property is being referenced anywhere in the call expression -
			// aliases will be required since the potentially error-trapped work will be performed in a lambda and "ref" arguments may not be accessed within a lambda. In this case,
			// there will be no need to overwrite the original reference with the alias after the call expression has been completed as the reference will not be passed to it ByRef
			// (if that was to be the case then it would have been picked up above).
			if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
			{
				foreach (var targetAsNameToken in callSetItemExpressionSegment.MemberAccessTokens.OfType<NameToken>())
				{
					if (!IsInByRefArgumentSet(byRefArguments, targetAsNameToken))
						continue;
					if (IsAlreadyAccountedFor(rewrittenReferences, targetAsNameToken, onlyConsiderNonReadOnlyMappings: false))
						continue;
					var rewrittenReferenceName = _tempNameGenerator(new CSharpName("byrefalias"), scopeAccessInformation);
					rewrittenReferences = rewrittenReferences.Add(new FuncByRefMapping(targetAsNameToken, rewrittenReferenceName, mappedValueIsReadOnly: true));
				}
			}

			foreach (var argument in callSetItemExpressionSegment.Arguments)
				rewrittenReferences = GetByRefArgumentsThatNeedRewriting(argument, scopeAccessInformation, rewrittenReferences, expressionIsDirectArgumentValue: true);

			return rewrittenReferences;
		}

		private bool IsInByRefArgumentSet(NonNullImmutableList<NameToken> byRefArguments, NameToken reference)
		{
			if (byRefArguments == null)
				throw new ArgumentNullException("byRefArguments");
			if (reference == null)
				throw new ArgumentNullException("reference");

			return byRefArguments.Any(
				byRefArgument => _nameRewriter.GetMemberAccessTokenName(byRefArgument) == _nameRewriter.GetMemberAccessTokenName(reference)
			);
		}

		/// <summary>
		/// Note: If onlyConsiderNonReadOnlyMappings is true then the mapping must have MappedValueIsReadOnly set to false. If onlyConsiderNonReadOnlyMappings is false then no
		/// filtering will be performed regarding the MappedValueIsReadOnly (it will not require that MappedValueIsReadOnly be true, it will not consider the value at all).
		/// </summary>
		private bool IsAlreadyAccountedFor(NonNullImmutableList<FuncByRefMapping> mappings, NameToken reference, bool onlyConsiderNonReadOnlyMappings)
		{
			if (mappings == null)
				throw new ArgumentNullException("mappings");
			if (reference == null)
				throw new ArgumentNullException("reference");

			return mappings.Any(mapping =>
				(_nameRewriter.GetMemberAccessTokenName(reference) == _nameRewriter.GetMemberAccessTokenName(mapping.From)) &&
				(!onlyConsiderNonReadOnlyMappings || !mapping.MappedValueIsReadOnly)
			);
		}

		private NonNullImmutableList<FuncByRefMapping> RemoveMappingForNameToken(NonNullImmutableList<FuncByRefMapping> mappings, NameToken reference)
		{
			if (mappings == null)
				throw new ArgumentNullException("mappings");
			if (reference == null)
				throw new ArgumentNullException("reference");

			return mappings
				.Where(mapping => _nameRewriter.GetMemberAccessTokenName(reference) == _nameRewriter.GetMemberAccessTokenName(mapping.From))
				.ToNonNullImmutableList();
		}

		private static NameToken TryToGetExpressionAsSingleNameToken(Expression expression)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");

			if (expression.Segments.Count() != 1)
				return null;

			var callExpressionSegment = expression.Segments.Single() as CallExpressionSegment;
			if ((callExpressionSegment == null)
			|| (callExpressionSegment.MemberAccessTokens.Count() != 1)
			|| callExpressionSegment.Arguments.Any()
			|| (callExpressionSegment.ZeroArgumentBracketsPresence != CallSetItemExpressionSegment.ArgumentBracketPresenceOptions.Absent))
				return null;

			return callExpressionSegment.MemberAccessTokens.Single() as NameToken;
		}
	}
}
