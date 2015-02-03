using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class DoBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public DoBlockTranslator(
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

		public TranslationResult Translate(DoBlock doBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (doBlock == null)
                throw new ArgumentNullException("doBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            TranslatedStatementContentDetails whileConditionExpressionContentIfAny;
            if (doBlock.ConditionIfAny == null)
                whileConditionExpressionContentIfAny = null;
            else
            {
                whileConditionExpressionContentIfAny = _statementTranslator.Translate(
                    doBlock.ConditionIfAny,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.Boolean,
                    _logger.Warning
                );
                if (!doBlock.IsDoWhileCondition)
                {
                    // C# doesn't support "DO UNTIL x" but it's equivalent to "DO WHILE !x"
                    whileConditionExpressionContentIfAny = new TranslatedStatementContentDetails(
                        "!" + whileConditionExpressionContentIfAny.TranslatedContent,
                        whileConditionExpressionContentIfAny.VariablesAccessed
                    );
                }
                if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
                {
                    // Ensure that the frankly ludicrous VBScript error-handling is applied where required. As the IProvideVBScriptCompatFunctionality's IF method
                    // signature describes, if an error occurs in retrieving the value, it will be evaulated as true. So, given a function
                    //
                    //   FUNCTION GetValue()
                    //     Err.Raise vbObjectError, "Test", "Test"
                    //   END FUNCTION
                    //
                    // both of the following loops will be entered:
                    //
                    //   ON ERROR RESUME NEXT
                    //   DO WHILE GetValue()
                    //     WScript.Echo "True"
                    //     EXIT DO
                    //   LOOP
                    //
                    //   ON ERROR RESUME NEXT
                    //   DO UNTIL GetValue()
                    //     WScript.Echo "True"
                    //     EXIT DO
                    //   LOOP
                    //
                    // This is why an additional IF call must be wrapped around the IF above - since the negation of a DO UNTIL (as opposed to a DO WHILE) loop
                    // must be within the outer, error-handling IF call. (A simpler example of the above is to replace the GetValue() call with 1/0, which will
                    // result in a "Division by zero" error if ON ERROR RESUME NEXT is not present, but which will result in both of the above loops being
                    // entered if it IS present).
                    whileConditionExpressionContentIfAny = new TranslatedStatementContentDetails(
                        string.Format(
                            "{0}.IF(() => {1}, {2})",
                            _supportRefName.Name,
                            whileConditionExpressionContentIfAny.TranslatedContent,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        whileConditionExpressionContentIfAny.VariablesAccessed
                    );
                }
            }

            var translationResult = TranslationResult.Empty;
            if (whileConditionExpressionContentIfAny != null)
            {
                translationResult = translationResult.AddUndeclaredVariables(
                    whileConditionExpressionContentIfAny.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            if (whileConditionExpressionContentIfAny == null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "while (true)",
                    indentationDepth
                ));
            }
            else if (doBlock.IsPreCondition)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "while (" + whileConditionExpressionContentIfAny.TranslatedContent + ")",
                    indentationDepth
                ));
            }
            else
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "do",
                    indentationDepth
                ));
            }
            translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
            var earlyExitNameIfAny = GetEarlyExitNameIfRequired(doBlock, scopeAccessInformation);
            if (earlyExitNameIfAny != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format("var {0} = false;", earlyExitNameIfAny.Name),
                    indentationDepth + 1
                ));
            }
            translationResult = translationResult.Add(
                Translate(doBlock.Statements.ToNonNullImmutableList(), scopeAccessInformation, doBlock.SupportsExit, earlyExitNameIfAny, indentationDepth + 1)
            );
            if ((whileConditionExpressionContentIfAny == null) || doBlock.IsPreCondition)
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            else
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    "} while (" + whileConditionExpressionContentIfAny.TranslatedContent + ");",
                    indentationDepth
                ));
            }
            var earlyExitFlagNamesToCheck = scopeAccessInformation.StructureExitPoints
                .Where(e => e.ExitEarlyBooleanNameIfAny != null)
                .Select(e => e.ExitEarlyBooleanNameIfAny.Name);
            if (earlyExitFlagNamesToCheck.Any())
            {
                // Perform early-exit checks for any scopeAccessInformation.StructureExitPoints - if this is DO..LOOP loop inside a FOR loop and an
                // EXIT FOR was encountered within the DO..LOOP that must refer to the containing FOR, then the DO..LOOP will have been broken out
                // of, but also a flag set that means that we must break further to get out of the FOR loop.
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        "if (" + string.Join(" || ", earlyExitFlagNamesToCheck) + ")",
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement(
                        "break;",
                        indentationDepth + 1
                    ));
            }
            return translationResult;
		}

        private CSharpName GetEarlyExitNameIfRequired(DoBlock doBlock, ScopeAccessInformation scopeAccessInformation)
        {
            if (doBlock == null)
                throw new ArgumentNullException("doBlock");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            if (!doBlock.SupportsExit || !doBlock.ContainsLoopThatContainsMismatchedExitThatMustBeHandledAtThisLevel())
                return null;

            return _tempNameGenerator(new CSharpName("exitDo"), scopeAccessInformation);
        }

		private TranslationResult Translate(
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeAccessInformation scopeAccessInformation,
            bool supportsExit,
            CSharpName earlyExitNameIfAny,
            int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (!supportsExit && (earlyExitNameIfAny != null))
                throw new ArgumentException("earlyExitNameIfAny should always be null if supportsExit is false");
            if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // Add a StructureExitPoint entry for the current loop so that the "early-exit" logic described in the Translate method above is possible
            if (supportsExit)
            {
                scopeAccessInformation = scopeAccessInformation.AddStructureExitPoints(
                    earlyExitNameIfAny,
                    ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions.Do
                );
            }
            return base.TranslateCommon(
                base.GetWithinFunctionBlockTranslators(),
				blocks,
                scopeAccessInformation,
                indentationDepth
			);
		}
    }
}
