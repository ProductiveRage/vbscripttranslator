using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation
{
    public abstract class CodeBlockTranslator
    {
		protected readonly CSharpName _supportClassName;
		protected readonly VBScriptNameRewriter _nameRewriter;
		protected readonly TempValueNameGenerator _tempNameGenerator;
		protected readonly ITranslateIndividualStatements _statementTranslator;
        protected CodeBlockTranslator(
            CSharpName supportClassName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator)
        {
            if (supportClassName == null)
                throw new ArgumentNullException("supportClassName");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (tempNameGenerator == null)
                throw new ArgumentNullException("tempNameGenerator");
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");

            _supportClassName = supportClassName;
            _nameRewriter = nameRewriter;
            _tempNameGenerator = tempNameGenerator;
            _statementTranslator = statementTranslator;
        }

		/// <summary>
		/// This should return null if it is unable to process the specified block. It should raise an exception for any null arguments. The returned value
        /// (where non-null) should overwrite the input translationResult in the caller's scope (it should not be added to it).
		/// </summary>
		protected delegate TranslationResult BlockTranslationAttempter(
			TranslationResult translationResult,
			ICodeBlock block,
			ScopeAccessInformation scopeAccessInformation,
			int indentationDepth
		);

		protected TranslationResult TranslateCommon(
			NonNullImmutableList<BlockTranslationAttempter> translators,
			NonNullImmutableList<ICodeBlock> blocks,
			ScopeAccessInformation scopeAccessInformation,
			int indentationDepth)
        {
			if (translators == null)
				throw new ArgumentNullException("translators");
            if (blocks == null)
                throw new ArgumentNullException("block");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			// TODO: Going to need to incorporate On Error Resume Next / Goto 0 handling outside of the rest of the process, requiring additional data
			// in the scopeAccessInformation type?

			// TODO: Ensure that the UndeclaredVariablesAccessed set is being 

            var translationResult = TranslationResult.Empty;
			foreach (var block in blocks)
            {
				var hasBlockBeenTranslated = false;
				foreach (var translator in translators)
				{
					var blockTranslationResult = translator(translationResult, block, scopeAccessInformation, indentationDepth);
					if (blockTranslationResult == null)
						continue;

                    translationResult = blockTranslationResult;
					hasBlockBeenTranslated = true;
                    break;
				}
				if (!hasBlockBeenTranslated)
					throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
            }

            // If the current parent construct doesn't affect scope (like IF and WHILE and unlike CLASS and FUNCTION) then the translationResult
            // can be returned directly and the nearest construct that does affect scope will be responsible for translating any explicit
            // variable declarations into translated statements
			if (!(scopeAccessInformation.ParentIfAny is IDefineScope))
                return translationResult;
            
            return FlushExplicitVariableDeclarations(
                translationResult,
                indentationDepth
            );
        }

		protected TranslationResult TryToTranslateBlankLine(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
            return (block is BlankLine) ? translationResult.Add(new TranslatedStatement("", indentationDepth)) : null;
		}

		protected TranslationResult TryToTranslateClass(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var classBlock = block as ClassBlock;
			if (classBlock == null)
				return null;

			var codeBlockTranslator = new ClassBlockTranslator(_supportClassName, _nameRewriter, _tempNameGenerator, _statementTranslator);
			return translationResult.Add(
				codeBlockTranslator.Translate(
					classBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateComment(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var commentBlock = block as CommentStatement;
			if (commentBlock == null)
				return null;

			var translatedCommentContent = "//" + commentBlock.Content;
			if (block is InlineCommentStatement)
			{
				var lastTranslatedStatement = translationResult.TranslatedStatements.LastOrDefault();
				if ((lastTranslatedStatement != null) && (lastTranslatedStatement.Content != ""))
				{
					translationResult = new TranslationResult(
						translationResult.TranslatedStatements
							.RemoveLast()
							.Add(new TranslatedStatement(
								lastTranslatedStatement.Content + " " + translatedCommentContent,
								lastTranslatedStatement.IndentationDepth
							)),
						translationResult.ExplicitVariableDeclarations,
						translationResult.UndeclaredVariablesAccessed
					);
					return translationResult;
				}
			}

			return translationResult.Add(
				new TranslatedStatement(translatedCommentContent, indentationDepth)
			);
		}

		protected TranslationResult TryToTranslateDim(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			// This covers the DimStatement, ReDimStatement, PrivateVariableStatement and PublicVariableStatement
			var explicitVariableDeclarationBlock = block as DimStatement;
			if (explicitVariableDeclarationBlock == null)
				return null;

			translationResult = translationResult.Add(
				explicitVariableDeclarationBlock.Variables.Select(v =>
					new VariableDeclaration(
						v.Name,
						(explicitVariableDeclarationBlock is PublicVariableStatement)
							? VariableDeclarationScopeOptions.Public
							: VariableDeclarationScopeOptions.Private,
						(v.Dimensions != null) // If the Dimensions set is non-null then this is an array type
					)
				)
			);

			var areDimensionsRequired = (
				(explicitVariableDeclarationBlock is ReDimStatement) ||
				(explicitVariableDeclarationBlock.Variables.Any(v => (v.Dimensions != null) && (v.Dimensions.Count > 0)))
			);
			if (!areDimensionsRequired)
				return translationResult;

			// TODO: If this is a ReDim then non-constant expressions may be used to set the dimension limits, in which case it may not be moved
			// (though a default-null declaration SHOULD be added as well as leaving the ReDim translation where it is)
			// TODO: Need a translated statement if setting dimensions
			throw new NotImplementedException("Not enabled support for declaring array variables with specifid dimensions yet");
		}

		protected TranslationResult TryToTranslateDo(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var doBlock = block as DoBlock;
			if (doBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateExit(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var exitStatement = block as ExitStatement;
			if (exitStatement == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateFor(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var forBlock = block as ForBlock;
			if (forBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateForEach(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var forEachBlock = block as ForEachBlock;
			if (forEachBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateFunction(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var functionBlock = block as AbstractFunctionBlock;
			if (functionBlock == null)
				return null;

			var codeBlockTranslator = new FunctionBlockTranslator(_supportClassName, _nameRewriter, _tempNameGenerator, _statementTranslator);
			return translationResult.Add(
				codeBlockTranslator.Translate(
					functionBlock,
					scopeAccessInformation,
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateIf(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var ifBlock = block as IfBlock;
			if (ifBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateOptionExplicit(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			return (block is OptionExplicit) ? TranslationResult.Empty : null;
		}

		protected TranslationResult TryToTranslateProperty(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var propertyBlock = block as PropertyBlock;
			if (propertyBlock == null)
				return null;

			// Note: Check for Default on the Get (the only place it's valid) before rendering Let or Set
			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateRandomize(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var randomizeStatement = block as RandomizeStatement;
			if (randomizeStatement == null)
				return null;

			// Note: See http://msdn.microsoft.com/en-us/library/e566zd96(v=vs.84).aspx when implementing RND
			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateSelect(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var selectBlock = block as SelectBlock;
			if (selectBlock == null)
				return null;

			throw new NotSupportedException(block.GetType() + " translation is not supported yet");
		}

		protected TranslationResult TryToTranslateStatementOrExpression(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			// This covers Statement and Expression instances as Expression inherits from Statement
			var statementBlock = block as Statement;
			if (statementBlock == null)
				return null;

			return translationResult.Add(
				new TranslatedStatement(
					_statementTranslator.Translate(statementBlock, scopeAccessInformation) + ";",
					indentationDepth
				)
			);
		}

		protected TranslationResult TryToTranslateValueSettingStatement(TranslationResult translationResult, ICodeBlock block, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			var valueSettingStatement = block as ValueSettingStatement;
			if (valueSettingStatement == null)
				return null;

            var assignmentFormat = GetAssignmentFormat(valueSettingStatement, scopeAccessInformation);

            var translatedExpression = _statementTranslator.Translate(
                valueSettingStatement.Expression,
                scopeAccessInformation,
                (valueSettingStatement.ValueSetType == ValueSettingStatement.ValueSetTypeOptions.Set)
                    ? ExpressionReturnTypeOptions.Reference
                    : ExpressionReturnTypeOptions.Value
            );

			// TODO: See http://blogs.msdn.com/b/ericlippert/archive/2004/08/19/error-handling-in-vbscript-part-one.aspx about value-setting statement
			// behaviour and On Error Resume Next - essentially, need to wrap the expression in try..catch as well as the setting, since if the expression
			// execution fails then the target should be set to empty (but this will fail if it's a SET assignment and so a further layer of error trapping
			// is required).
			return translationResult.Add(
                new TranslatedStatement(
                    assignmentFormat(translatedExpression) + ";",
                    indentationDepth
                )
            );
		}

        private Func<string, string> GetAssignmentFormat(ValueSettingStatement valueSettingStatement, ScopeAccessInformation scopeAccessInformation)
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

                // TODO: Need to consider whether we're in a function or property and whether this is an assigment for its return value
                return translatedExpression => string.Format(
                    "{0} = {1}",
                    _nameRewriter(singleTokenAsName).Name,
                    translatedExpression
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
                        new VBScriptTranslator.StageTwoParser.ExpressionParsing.Expression[0]
                    )
                );
                callExpressionSegments.Add(
                    new CallExpressionSegment(
                        new[] { lastCallExpressionSegments.MemberAccessTokens.Last() },
                        lastCallExpressionSegments.Arguments
                    )
                );
            }

            // Now we either have a single segment in the set that has no more than two member accessors (lending itself to easy extraction of a
            // target and optional member accessor) or we have multiple segments where all but the last one define the target and the last entry
            // has the optional member accessor and any arguments.
            string targetAccessor;
            string optionalMemberAccessor;
            IEnumerable<VBScriptTranslator.StageTwoParser.ExpressionParsing.Expression> arguments;
            if (callExpressionSegments.Count == 1)
            {
                // The single CallExpressionSegment may have one or two member accessors
                targetAccessor = callExpressionSegments[0].MemberAccessTokens.First().Content;
                if (callExpressionSegments[0].MemberAccessTokens.Count() == 1)
                    optionalMemberAccessor = null;
                else
                    optionalMemberAccessor = callExpressionSegments[0].MemberAccessTokens.Skip(1).Single().Content;
                arguments = callExpressionSegments[0].Arguments;
            }
            else
            {
                var targetAccessCallExpressionSegments = callExpressionSegments.Take(callExpressionSegments.Count() - 1);
                var targetAccessExpressionSegments = (targetAccessCallExpressionSegments.Count() > 1)
                    ? new IExpressionSegment[] { new CallSetExpressionSegment(targetAccessCallExpressionSegments) }
                    : new IExpressionSegment[] { targetAccessCallExpressionSegments.Single() };
                targetAccessor = _statementTranslator.Translate(
                    new VBScriptTranslator.StageTwoParser.ExpressionParsing.Expression(
                        targetAccessExpressionSegments
                    ),
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified
                );

                // The last CallExpressionSegment may only have one member accessor
                var lastCallExpressionSegment = callExpressionSegments.Last();
                optionalMemberAccessor = lastCallExpressionSegment.MemberAccessTokens.Single().Content;
                arguments = lastCallExpressionSegment.Arguments;
            }

            // Note: The translatedExpression will already account for whether the statement is of type LET or SET
            return translatedExpression => string.Format(
                "{0}.SET({1}, {2}, {3}, {4})",
                _supportClassName.Name,
                targetAccessor,
                (optionalMemberAccessor == null) ? "null" : optionalMemberAccessor.ToLiteral(),
				arguments.Any()
					? string.Format(
						"new object[] {{ {0} }}",
						string.Join(
							", ",
							arguments.Select(a => _statementTranslator.Translate(a, scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified))
						)
					)
					: "new object[0]",
                translatedExpression
            );
        }

        protected string TranslateVariableDeclaration(VariableDeclaration variableDeclaration)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");

            return string.Format(
                "object {0} = {1}null;",
                _nameRewriter(variableDeclaration.Name).Name,
                variableDeclaration.IsArray ? "(object[])" : ""
            );
        }

		protected TranslationResult FlushExplicitVariableDeclarations(TranslationResult translationResult, int indentationDepth)
		{
			// TODO: Consider trying to insert the content after any comments or blank lines?
			if (translationResult == null)
				throw new ArgumentNullException("translationResult");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			return new TranslationResult(
				translationResult.ExplicitVariableDeclarations
					.Select(v =>
						 new TranslatedStatement(TranslateVariableDeclaration(v), indentationDepth)
					)
					.ToNonNullImmutableList()
					.AddRange(translationResult.TranslatedStatements),
				new NonNullImmutableList<VariableDeclaration>(),
				translationResult.UndeclaredVariablesAccessed
			);
		}
    }
}
