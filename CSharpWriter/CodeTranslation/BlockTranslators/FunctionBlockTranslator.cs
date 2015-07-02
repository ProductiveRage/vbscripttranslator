using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSharpSupport.Attributes;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class FunctionBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public FunctionBlockTranslator(
            CSharpName supportRefName,
            CSharpName envClassName,
            CSharpName envRefName,
            CSharpName outerClassName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
			ITranslateIndividualStatements statementTranslator,
			ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
            : base(supportRefName, envClassName, envRefName, outerClassName, outerRefName, nameRewriter, tempNameGenerator, statementTranslator, valueSettingStatementTranslator, logger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException("statementTranslator");
            if (logger == null)
                throw new ArgumentNullException("logger");

            _statementTranslator = statementTranslator;
            _logger = logger;
        }

		public TranslationResult Translate(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (functionBlock == null)
				throw new ArgumentNullException("functionBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var isSingleReturnValueStatementFunction = IsSingleReturnValueStatementFunctionWithoutAnyByRefMappings(functionBlock, scopeAccessInformation);
            var returnValueName = functionBlock.HasReturnValue ? _tempNameGenerator(new CSharpName("retVal"), scopeAccessInformation) : null;
			var translationResult = TranslationResult.Empty.Add(
				TranslateFunctionHeader(
					functionBlock,
                    scopeAccessInformation,
					returnValueName,
					indentationDepth
				)
			);
            CSharpName errorRegistrationTokenIfAny;
            if (functionBlock.Statements.ToNonNullImmutableList().DoesScopeContainOnErrorResumeNext())
            {
                errorRegistrationTokenIfAny = _tempNameGenerator(new CSharpName("errOn"), scopeAccessInformation);
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(
                        "var {0} = {1}.GETERRORTRAPPINGTOKEN();",
                        errorRegistrationTokenIfAny.Name,
                        _supportRefName.Name
                    ),
                    indentationDepth + 1
                ));
            }
            else
                errorRegistrationTokenIfAny = null;
            translationResult = translationResult.Add(
				Translate(
				    functionBlock.Statements.ToNonNullImmutableList(),
					    scopeAccessInformation.Extend(
                            functionBlock,
                            returnValueName,
                            errorRegistrationTokenIfAny,
                            functionBlock.Statements.ToNonNullImmutableList()
                        ),
                    isSingleReturnValueStatementFunction,
				    indentationDepth + 1
				)
			);
            if (errorRegistrationTokenIfAny != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(
                        "{0}.RELEASEERRORTRAPPINGTOKEN({1});",
                        _supportRefName.Name,
                        errorRegistrationTokenIfAny.Name
                    ),
                    indentationDepth + 1
                ));
            }
            if (functionBlock.HasReturnValue && !isSingleReturnValueStatementFunction)
			{
				// If this is an empty function then just render "return null" (TranslateFunctionHeader won't declare the return value reference) 
				translationResult = translationResult
                    .Add(new TranslatedStatement(
						string.Format(
							"return {0};",
							functionBlock.Statements.Any() ? returnValueName.Name : "null"
						),
						indentationDepth + 1
					));
            }
			return translationResult.Add(
				new TranslatedStatement("}", indentationDepth)
            );
		}

		private TranslationResult Translate(
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeAccessInformation scopeAccessInformation,
            bool isSingleReturnValueStatementFunction,
            int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            NonNullImmutableList<BlockTranslationAttempter> blockTranslators;
            if (isSingleReturnValueStatementFunction)
            {
                blockTranslators = new NonNullImmutableList<BlockTranslationAttempter>()
                    .Add(TryToTranslateValueSettingStatementAsSimpleFunctionValueReturner)
                    .Add(base.TryToTranslateBlankLine)
                    .Add(base.TryToTranslateComment);
            }
            else
                blockTranslators = base.GetWithinFunctionBlockTranslators();

			return base.TranslateCommon(
                blockTranslators,
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
		}

        private TranslationResult TryToTranslateValueSettingStatementAsSimpleFunctionValueReturner(
            TranslationResult translationResult,
            ICodeBlock block,
            ScopeAccessInformation scopeAccessInformation,
            int indentationDepth)
        {
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (block == null)
                throw new ArgumentNullException("block");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var valueSettingStatement = block as ValueSettingStatement;
            if (valueSettingStatement == null)
                return null;

            var translatedStatementContentDetails = _statementTranslator.Translate(
                valueSettingStatement.Expression,
                scopeAccessInformation,
                (valueSettingStatement.ValueSetType == ValueSettingStatement.ValueSetTypeOptions.Set)
                    ? ExpressionReturnTypeOptions.Reference
                    : ExpressionReturnTypeOptions.Value,
                _logger.Warning
            );
            var undeclaredVariables = translatedStatementContentDetails.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(v, _nameRewriter));
            foreach (var undeclaredVariable in undeclaredVariables)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

            return translationResult
                .Add(new TranslatedStatement(
                    "return " + translatedStatementContentDetails.TranslatedContent + ";",
                    indentationDepth
                ))
    			.AddUndeclaredVariables(undeclaredVariables);
        }

        private IEnumerable<TranslatedStatement> TranslateFunctionHeader(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation, CSharpName returnValueNameIfAny, int indentationDepth)
		{
			if (functionBlock == null)
				throw new ArgumentNullException("functionBlock");
			if (functionBlock.HasReturnValue && (returnValueNameIfAny == null))
				throw new ArgumentException("returnValueNameIfAny must not be null if functionBlock.HasReturnValue is true");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			var content = new StringBuilder();
			content.Append(functionBlock.IsPublic ? "public" : "private");
			content.Append(" ");
			content.Append(functionBlock.HasReturnValue ? "object" : "void");
			content.Append(" ");
			content.Append(_nameRewriter.GetMemberAccessTokenName(functionBlock.Name));
			content.Append("(");
			var numberOfParameters = functionBlock.Parameters.Count();
			for (var index = 0; index < numberOfParameters; index++)
			{
				var parameter = functionBlock.Parameters.ElementAt(index);
				if (parameter.ByRef)
					content.Append("ref ");
                content.Append("object ");
                content.Append(_nameRewriter.GetMemberAccessTokenName(parameter.Name));
                if (index < (numberOfParameters - 1))
					content.Append(", ");
			}
			content.Append(")");

			var translatedStatements = new List<TranslatedStatement>();
			if (functionBlock.IsDefault)
				translatedStatements.Add(new TranslatedStatement("[" + typeof(IsDefault).FullName + "]", indentationDepth));
            var property = functionBlock as PropertyBlock;
            if ((property != null) && property.IsPublic && property.IsIndexedProperty())
            {
                translatedStatements.Add(
                    new TranslatedStatement(
                        string.Format(
                            "[TranslatedProperty({0})]", // Note: Safe to assume that using statements are present for the namespace that contains TranslatedProperty
                            property.Name.Content.ToLiteral()
                        ),
                        indentationDepth
                    )
                );
            }
            translatedStatements.Add(new TranslatedStatement(content.ToString(), indentationDepth));
            translatedStatements.Add(new TranslatedStatement("{", indentationDepth));
            if (functionBlock.HasReturnValue && functionBlock.Statements.Any() && !IsSingleReturnValueStatementFunctionWithoutAnyByRefMappings(functionBlock, scopeAccessInformation))
			{
				translatedStatements.Add(new TranslatedStatement(
					base.TranslateVariableInitialisation(
						new VariableDeclaration(
							new DoNotRenameNameToken(
                                returnValueNameIfAny.Name,
                                functionBlock.Name.LineIndex
                            ),
							VariableDeclarationScopeOptions.Private,
                            null // Not declared as an array
                        ),
                        ScopeLocationOptions.WithinFunctionOrPropertyOrWith
					),
					indentationDepth + 1
				));
			}
            return translatedStatements;
		}

        /// <summary>
        /// If a function or property only contains a single executable block, which is a return statement, then this can be translated into a simple return
        /// statement in the C# output (as opposed to having to maintain a temporary variable for the return value in case there are various manipulations
        /// of it or error-handling or any other VBScript oddnes required)
        /// </summary>
        private bool IsSingleReturnValueStatementFunctionWithoutAnyByRefMappings(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation)
        {
            if (functionBlock == null)
                throw new ArgumentNullException("functionBlock");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            var executableStatements = functionBlock.Statements.Where(s => !(s is INonExecutableCodeBlock));
            if (executableStatements.Count() != 1)
                return false;

            var valueSettingStatement = executableStatements.Single() as ValueSettingStatement;
            if (valueSettingStatement == null)
                return false;

            if (valueSettingStatement.ValueToSet.Tokens.Count() != 1)
                return false;

            var valueToSetTokenAsNameToken = valueSettingStatement.ValueToSet.Tokens.Single() as NameToken;
            if (valueToSetTokenAsNameToken == null)
                return false;

            if (_nameRewriter.GetMemberAccessTokenName(valueToSetTokenAsNameToken) != _nameRewriter.GetMemberAccessTokenName(functionBlock.Name))
                return false;

            // If there is no return value (ie. it's a SUB or a LET/SET PROPERTY accessor) then this can't apply (not only can this simple single-line
            // return format not be used but a runtime error is required if the value-setting statement targets the name of a SUB)
            if (!functionBlock.HasReturnValue)
                return false;

            // If any values need aliasing in order to perform this "one liner" then it won't be possible to represent it a simple one-line return, it will
            // need a try..finally setting up to create the alias(es), use where required and then map the values back over the original(s).
            scopeAccessInformation = scopeAccessInformation.Extend(functionBlock, functionBlock.Statements.ToNonNullImmutableList());
            var byRefArgumentMapper = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
            var byRefArgumentsToMap = byRefArgumentMapper.GetByRefArgumentsThatNeedRewriting(
                valueSettingStatement.Expression.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
                scopeAccessInformation,
                new NonNullImmutableList<FuncByRefMapping>()
            );
            if (byRefArgumentsToMap.Any())
                return false;

            return !valueSettingStatement.Expression.Tokens.Any(
                t => (t is NameToken) && (_nameRewriter.GetMemberAccessTokenName(t) == _nameRewriter.GetMemberAccessTokenName(functionBlock.Name))
            );
        }
    }
}
