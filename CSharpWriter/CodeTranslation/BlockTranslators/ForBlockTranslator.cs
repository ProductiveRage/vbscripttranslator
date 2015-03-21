using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using StageTwoExpressionParsing = VBScriptTranslator.StageTwoParser.ExpressionParsing;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class ForBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public ForBlockTranslator(
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

		public TranslationResult Translate(ForBlock forBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (forBlock == null)
                throw new ArgumentNullException("forBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // A note about ON ERROR RESUME NEXT and the intriguing behaviour it can introduce. The following will display "We're in the loop! i is Empty" -
            //
            //   On Error Resume Next
            //   Dim i: For i = 1 To 1/0
            //     WScript.Echo "We're in the loop! i is " & TypeName(i)
            //   Next
            //
            // If error-trapping is enabled and one of the From, To or Step evaluation fails, no further constaints will receive evaluation attempts but
            // the loop WILL be entered.. once. Only once. And the loop variable will not be initialised when doing so (so, in the above example, it has
            // the "Empty" value inside the loop - but if it had been set to "a" before the loop construct then it would still have value "a" within
            // the loop).
            // - Note: In http://blogs.msdn.com/b/ericlippert/archive/2004/08/19/error-handling-in-vbscript-part-one.aspx, If Blah raises an error then it resumes on the Print "Hello" in either case. 
            // - Note: In http://blogs.msdn.com/b/ericlippert/archive/2004/08/19/error-handling-in-vbscript-part-one.aspx, Eric Lippert does describe
            //   this (see the note that reads "If Blah raises an error, this resumes into the loop, not after the loop")

            // Identify tokens for the start, end and step variables. If they are numeric constants then use them, otherwise they must be stored in temporary
            // values. Note that these temporary values are NOT re-evaluated each loop since this is how VBScript (unlike some other languages) work. Also
            // note the loopStart is generated after loopEnd and loopStep - this is because it may depend upon them. In "FOR i = 1 TO 10", loopStart will be
            // a VBScript "Integer" (Int16) but in "FOR i = 1 TO 32768" loopStart needs to be a VBScript "Long" (Int32) since VBScript tries to arrange for
            // the loop variable to be of a type that doesn't need to change each iteration. (It's possible for code within the loop to change the loop
            // variable's type, but that's another story).
            // - Note: VBScript does not seem willing to consider the types Boolean or Byte to be elligible for use as a loop variable. This maybe isn't
            //   THAT surprising for Boolean but it seems strange for Byte. Regardless, the loop variable "i" in both examples is of type "Integer":
            //     For i = CBool(0) TO CBool(1)
            //     For i = CByte(0) TO CByte(1) ' Even though CByte returns type "Byte", the loop variable "i" here is always "Integer"
            var undeclaredVariableReferencesAccessedByLoopConstraints = new NonNullImmutableList<NameToken>();
            var loopConstraintInitialisersWhereRequired = new List<LoopConstraintInitialiser>();
            string loopEnd;
            var numericLoopEndValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopTo);
            if (numericLoopEndValueIfAny != null)
            {
                // Note: We need to use the NumericValueToken's AsCSharpValue() method here when generating the translated "loopEnd" output since its
                // type might be important when determining what the type for "loopStart" will be - eg. the end value in "FOR i = 1 TO 20.2" results
                // in the loop variable being a double.
                loopEnd = numericLoopEndValueIfAny.AsCSharpValue();
            }
            else
            {
                var numericLoopEndContent = WrapInNUMCallIfRequired(forBlock.LoopTo, scopeAccessInformation);
                var loopEndName = _tempNameGenerator(new CSharpName("loopEnd"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(new LoopConstraintInitialiser(
                    loopEndName,
                    numericLoopEndContent.TranslatedContent
                ));
                loopEnd = loopEndName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    numericLoopEndContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }
            string loopStep;
            var numericLoopStepValueIfAny = (forBlock.LoopStep == null)
                ? new NumericValueToken("1", forBlock.LoopTo.Tokens.Last().LineIndex) // Default to Step 1 if no LoopStep expression was specified
                : TryToGetExpressionAsNumericConstant(forBlock.LoopStep);
            if ((numericLoopStepValueIfAny == null) && (forBlock.LoopStep.Tokens.Count() == 2))
            {
                // If the loop step is "-1" then this will still appear as distinct tokens "-" and "1" since the NumberRebuilder can't safely combine
                // the two tokens at the level of knowledge that it has, since it doesn't know if the "STEP" token is a keyword (meaning it's part of
                // a loop structure and so they could be safely combined) or a reference name (meaning it could be something like "step - 1", in which
                // case they could not be safely combined).
                var firstLoopStepToken = forBlock.LoopStep.Tokens.First();
                var lastLoopStepTokenAsNumericValueToken = forBlock.LoopStep.Tokens.Last() as NumericValueToken;
                if ((firstLoopStepToken is OperatorToken) && (firstLoopStepToken.Content == "-") && (lastLoopStepTokenAsNumericValueToken != null))
                    numericLoopStepValueIfAny = lastLoopStepTokenAsNumericValueToken.GetNegative();
            }
            if (numericLoopStepValueIfAny != null)
            {
                // Note: We need to use the NumericValueToken's AsCSharpValue() method here when generating the translated "loopStep" output since its
                // type might be important when determining what the type for "loopStart" will be - eg. the loop step in "FOR i = 1 TO 10 STEP 0.1"
                // results in the loop variable being a double.
                loopStep = numericLoopStepValueIfAny.AsCSharpValue();
            }
            else
            {
                var numericLoopStepContent = WrapInNUMCallIfRequired(forBlock.LoopStep, scopeAccessInformation);
                var loopStepName = _tempNameGenerator(new CSharpName("loopStep"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(new LoopConstraintInitialiser(
                    loopStepName,
                    numericLoopStepContent.TranslatedContent
                ));
                loopStep = loopStepName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    numericLoopStepContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            // The loopStart value is what the loop variable will first be set to. This means that it has to follow some complicate VBScript logic. The
            // VBScript interpreter tries to ensure that a loop variable's type will not change during an iteration (unless something crazy within in
            // loop changes it, which is allowed to happen - but not something we need to worry about here). So in the loop "FOR i = 1 To 5" it's fine
            // for the loopStart to be a VBScript "Integer" (Int16) since that can run the loop from 1..5 without changing type. However, in the loop
            // "FOR i = 1 To 32768" an Int16 wouldn't be able to describe all of the values, 32768 would be an overflow. So VBScript would set the
            // loop variable to be a "Long" (Int32). It's important that the loopStart expression we generate here includes sufficient type information
            // to do the same sort of thing. (There is similar-but-different handling for dates - eg. "FOR i = Date() TO Date() + 2" - where the type
            // of the loop variable will be a Date, but if the loopEnd expression is too high - eg. "FOR i = Date() To 10000000" - then there will be
            // an overflow error, rather than the loop variable type being changed so that it can contain the entire range).
            string loopStart;
            var numericLoopStartValueIfAny = TryToGetExpressionAsNumericConstant(forBlock.LoopFrom);
            if ((numericLoopStartValueIfAny != null) && numericLoopStartValueIfAny.IsVBScriptInteger()
            && (numericLoopEndValueIfAny != null) && numericLoopEndValueIfAny.IsVBScriptInteger()
            && (numericLoopStepValueIfAny != null) && numericLoopStepValueIfAny.IsVBScriptInteger())
            {
                // For really simple loops, like "FOR i = 1 TO 5", where the loop constraints are all compile-time constants and all within the "Integer"
                // range, bypass the logic that performs the other checks and just call numericLoopStartValueIfAny.AsCSharpValue() - this will include
                // type information in the translated content (eg. "(Int16)1").
                // - One the one hand, it seems a bit silly bothering layering on more special cases when VBScript already has a million special cases
                //   of its own, but the exception here makes the translated output slightly less WTF-worthy for simple and common cases
                loopStart = numericLoopStartValueIfAny.AsCSharpValue();
            }
            else
            {
                // Note: Previously assumed that if the loop end and/or step values were known integer constants that they could be ignored when determining
                // the loop variable. This is incorrect since "FOR i = CBYTE(1) TO CBYTE(5)" results in the loop variable "i" being an "Integer" since the
                // implicit step is of type "Integer", in order for "i" to be of type "Byte" the loop must be "FOR i = CBYTE(1) TO CBYTE(5) STEP CBYTE(1)".
                // However, there is one a minor shortcut we can take, don't include duplicate values in the NUM call - so if we have "FOR i = 1 To a", the
                // loop start and step are the same, so instead of emitting "NUM((Int16)1, a, (Int16)1)" trim it down to "NUM((Int16)1, a)".
                var loopStartExpressionContent = _statementTranslator.Translate(
                    forBlock.LoopFrom,
                    scopeAccessInformation,
                    ExpressionReturnTypeOptions.NotSpecified,
                    _logger.Warning
                );
                var numericValuesTheTypeMustBeAbleToContain = new List<string>();
                if (loopEnd != loopStartExpressionContent.TranslatedContent)
                    numericValuesTheTypeMustBeAbleToContain.Add(loopEnd);
                if ((loopStep != loopStartExpressionContent.TranslatedContent) && (loopStep != loopEnd))
                    numericValuesTheTypeMustBeAbleToContain.Add(loopStep);

                // The LoopStartConstraintInitialiser takes both two "initialisation content" parameters - one to initialise its content without taking
                // into account the other constraints and one that DOES take into account the others. This will be important further down since if a
                // loop's constraints are all individually evaluated ok and error-trapping is enabled and the loop start value is a Date but one of
                // the others would cause an Overflow, the loop will be entered once with the loop variable set to the loop start Date value. This
                // is different to the case where error-trapping is enabled and one of the loop constraint evaluation fails; in that case, the
                // loop will still be entered once but the loop variable will not be set.
                var loopStartName = _tempNameGenerator(new CSharpName("loopStart"), scopeAccessInformation);
                loopConstraintInitialisersWhereRequired.Add(new LoopStartConstraintInitialiser(
                    loopStartName,
                    string.Format(
                        "{0}.NUM({1})", // This is the format of the content where other types are never taken into account
                        _supportRefName.Name,
                        loopStartExpressionContent.TranslatedContent
                    ),
                    string.Format(
                        "{0}.NUM({1}{2}{3})", // This is the initialisation content where other types will be taken into account, where relevant
                        _supportRefName.Name,
                        loopStartExpressionContent.TranslatedContent,
                        numericValuesTheTypeMustBeAbleToContain.Any() ? ", " : "",
                        string.Join(", ", numericValuesTheTypeMustBeAbleToContain)
                    )
                ));
                loopStart = loopStartName.Name;
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.AddRange(
                    loopStartExpressionContent.GetUndeclaredVariablesAccessed(scopeAccessInformation, _nameRewriter)
                );
            }

            var rewrittenLoopVariableName = _nameRewriter.GetMemberAccessTokenName(forBlock.LoopVar);
            var targetContainer = scopeAccessInformation.GetNameOfTargetContainerIfAnyRequired(rewrittenLoopVariableName, _envRefName, _outerRefName, _nameRewriter);
            if (targetContainer != null)
                rewrittenLoopVariableName = targetContainer.Name + "." + rewrittenLoopVariableName;
            if (!scopeAccessInformation.IsDeclaredReference(rewrittenLoopVariableName, _nameRewriter))
                undeclaredVariableReferencesAccessedByLoopConstraints = undeclaredVariableReferencesAccessedByLoopConstraints.Add(forBlock.LoopVar);

            // Any dynamic loop constraints (ie. those that can be confirmed to be fixed numeric values at translation time) need to have variables
            // declared and initialised. There may be a mix of dynamic and constant constraints so there may be zero, one, two or three variables
            // to deal with here (this will have been determined in the work above).
            var translationResult = TranslationResult.Empty;
            CSharpName constraintsInitialisedFlagNameIfAny;
            if (!loopConstraintInitialisersWhereRequired.Any())
                constraintsInitialisedFlagNameIfAny = null;
            else if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
            {
                constraintsInitialisedFlagNameIfAny = null;
                foreach (var loopConstraintInitialiser in loopConstraintInitialisersWhereRequired)
                {
                    translationResult = translationResult
                        .Add(new TranslatedStatement(
                            string.Format(
                                "var {0} = {1};",
                                loopConstraintInitialiser.VariableName.Name,
                                loopConstraintInitialiser.InitialisationContent
                            ),
                            indentationDepth
                        ));
                }
            }
            else
            {
                if (loopConstraintInitialisersWhereRequired.Any())
                {
                    translationResult = translationResult.Add(new TranslatedStatement(
                        "object " + string.Join(", ", loopConstraintInitialisersWhereRequired.Select(l => l.VariableName.Name + " = 0")) + ";",
                        indentationDepth
                    ));
                }
                constraintsInitialisedFlagNameIfAny = _tempNameGenerator(new CSharpName("loopConstraintsInitialised"), scopeAccessInformation);
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "var {0} = false;",
                            constraintsInitialisedFlagNameIfAny.Name
                        ),
                        indentationDepth
                    ))
                    .Add(new TranslatedStatement(
                        string.Format(
                            "{0}.HANDLEERROR({1}, () => {{",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ));
                LoopStartConstraintInitialiser loopStartConstraintInitialiserIfAny = null;
                foreach (var loopConstraintInitialiser in loopConstraintInitialisersWhereRequired)
                {
                    // Error-trapping may be enabled when this emitted code is being executed, so we need to consider the special case with Date loop variables
                    // and how they may overflow, which is why the start constraint's InitialisationContentIgnoringTypesOfOtherConstraints content is used here.
                    // When there are no dates to worry about, and error-trapping is enabled, then either all loop constraints evaluate fine and the loop executes
                    // as normal or one of the loop constraints fails to evaluate and the loop is (bizarrely) executed once only, without setting the loop variable.
                    // However, if the individual constraints all evaulate ok and error-trapping is enabled but the loop start value is a Date while another value
                    // is a numeric value that would overflow Date, then the loop is entered once and the loop variable is set to that start date. To achieve this,
                    // we try to initialise loopStart without considering loopEnd or loopStep - if this fails then it's the standard loop-constraint-evaluation-fail
                    // and so the loop variable will not be set and the loop will be executed once only. If, however, loopStart is evaluated successfully in isolation,
                    // then the loop variable is set to then and then loopStart will be re-evaluated with consideration to loopEnd and loopStart - if this fails, then
                    // the loop-constraints-initialised flag is still false and so the loop will be executed only once BUT this time the loop variable is set to the
                    // first evaluation of loopStart. This works because a Date is the only value type that may be evaluated as a number in isolation but fail when
                    // loopEnd and loopStep are considered (if loopStart was an Int16 and loopEnd an Int32 then the loop variable type will be expanded to fit the
                    // entire range, but this is not possible if loopStart is a DateTime and loopEnd a Double outside of the VBScript Date range since there is
                    // no "bigger" type that is still a date for the loop variable to use). These lines are all processed within a HANDLEERROR call, so when one of
                    // then fails, the subsequent lines will not be called - this is how the try-to-set-loopStart logic twice works.
                    string initialisationContentToUse;
                    if (loopConstraintInitialiser is LoopStartConstraintInitialiser)
                    {
                        if (loopStartConstraintInitialiserIfAny != null)
                            throw new Exception("The loopStartConstraintInitialisers set shouldn't have more than one LoopStartConstraintInitialiser!");
                        loopStartConstraintInitialiserIfAny = (LoopStartConstraintInitialiser)loopConstraintInitialiser;
                        initialisationContentToUse = loopStartConstraintInitialiserIfAny.InitialisationContentIgnoringTypesOfOtherConstraints;
                    }
                    else
                        initialisationContentToUse = loopConstraintInitialiser.InitialisationContent;
                    translationResult = translationResult
                        .Add(new TranslatedStatement(
                            string.Format(
                                "{0} = {1};",
                                loopConstraintInitialiser.VariableName.Name,
                                initialisationContentToUse
                            ),
                            indentationDepth + 1
                        ));
                }
                if (loopStartConstraintInitialiserIfAny != null)
                {
                    // This is logic explained above - where the loopStart value may be set twice to deal with overflow errors with loops with a Date-type loop variable.
                    // Note: If the loop-start initialisation does not depend upon loopEnd or loopStep then the InitialisationContentIgnoringTypesOfOtherConstraints value
                    // will be the same as loopStartConstraintInitialiserIfAny and there's no additional work to do (this may be the case if loopEnd and loopStart are both
                    // constants of type Int16 since they wouldn't have any effect on the loop variable).
                    if (loopStartConstraintInitialiserIfAny.InitialisationContent != loopStartConstraintInitialiserIfAny.InitialisationContentIgnoringTypesOfOtherConstraints)
                    {
                        // Actually, we ONLY set the loop variable to the loopStart value that was determined before considering loopEnd and loopStep if the loopStart is
                        // found to be a DateTime - this is so that the following two examples work (note that in both cases it is assumed that ON ERROR RESUME NEXT is
                        // in play since otherwise the loops won't be entered at all since there is an overflow error):
                        //   FOR i = Date() To 10000000 ' The loop will be entered once, "i" will be set to Date()
                        //   FOR i = 10000000 To Date() ' The loop will be entered once, "i" will NOT be set
                        // Update: The same logic applies to Decimal (Currency in VBScript) since this is the only other type that VBScript sticks to even when it needs
                        // to move up to a larger data type (such as Double). The below two examples are equivalent to the above and exhibit the same behaviour (note
                        // that the max value for a Currency is  922,337,203,685,477.5807 - see https://msdn.microsoft.com/en-us/library/9e7a57cf%28v=vs.84%29.aspx)
                        //   FOR i = CCur(922337203685475) TO CCur(922337203685476) STEP CDBL("9223372036854760") ' Loop is entered once, "i" is set to a Currency value
                        //   FOR i = CDbl("9223372036854760") TO CCur(922337203685475) STEP -1 ' Loop is entered once but "i" will not be set
                        translationResult = translationResult.Add(new TranslatedStatement(
                            string.Format(
                                "if (({0} is DateTime) || ({0} is Decimal))",
                                loopStart
                            ),
                            indentationDepth + 1
                        ));
                        translationResult = translationResult.Add(new TranslatedStatement(
                            string.Format(
                                "{0} = {1};",
                                rewrittenLoopVariableName,
                                loopStart
                            ),
                            indentationDepth + 2
                        ));
                        translationResult = translationResult.Add(new TranslatedStatement(
                            string.Format(
                                "{0} = {1};",
                                loopStart,
                                loopStartConstraintInitialiserIfAny.InitialisationContent
                            ),
                            indentationDepth + 1
                        ));
                    }
                }
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        constraintsInitialisedFlagNameIfAny.Name + " = true;",
                        indentationDepth + 1
                    ))
                    .Add(new TranslatedStatement("});", indentationDepth));
            }

            // If all three constraints are numeric constraints then we can determine now whether the loop should be executed or not
            // - If the loop is descending but a negative step is not specified then the loop will not be entered and so might as well not be
            //   included in the translated content
            // - If the loop is ascending then a positive OR zero step must be specified
            var guardClauseLines = new NonNullImmutableList<string>();
            if ((numericLoopStartValueIfAny != null) && (numericLoopEndValueIfAny != null) && (numericLoopStepValueIfAny != null))
            {
                if ((numericLoopStartValueIfAny.Value > numericLoopEndValueIfAny.Value) && (numericLoopStepValueIfAny.Value >= 0))
                {
                    _logger.Warning("Optimising out descending FOR loop that does not have a negative step specified");
                    return TranslationResult.Empty;
                }
                if ((numericLoopStartValueIfAny.Value <= numericLoopEndValueIfAny.Value) && (numericLoopStepValueIfAny.Value < 0))
                {
                    _logger.Warning("Optimising out ascending FOR loop that has a negative step specified");
                    return TranslationResult.Empty;
                }
            }
            else if (numericLoopStepValueIfAny != null)
            {
                // If the step is a known numeric constant, then the guard clause only needs to compare the from and to constaints (at least
                // one of which must not be a known constant, otherwise we'd be in the above condition)
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    // Ascending loop or infinite loop (step zero, which is supported in VBScript), start must not be greater than end
                    guardClauseLines = guardClauseLines.Add(string.Format("({0}.StrictLTE({1}, {2}))", _supportRefName.Name, loopStart, loopEnd));
                }
                else
                {
                    // Descending loop, start must be greater than end
                    guardClauseLines = guardClauseLines.Add(string.Format("({0}.StrictGT({1}, {2}))", _supportRefName.Name, loopStart, loopEnd));
                }
            }
            else if ((numericLoopStartValueIfAny != null) && (numericLoopEndValueIfAny != null))
            {
                // If the from and to bounds are known to be constants, but not the step then we just need to ensure that the step matches the
                // direction of from and to at runtime. Note: A step of zero will cause an infinite loop, but only if from <= to (the loop will
                // not be executed if it is descending and has a step of zero)
                if (numericLoopStartValueIfAny.Value <= numericLoopEndValueIfAny.Value)
                    guardClauseLines = guardClauseLines.Add(string.Format("{0}.StrictGTE({1}, 0)", _supportRefName.Name, loopStep));
                else
                    guardClauseLines = guardClauseLines.Add(string.Format("{0}.StrictLT({1}. 0)", _supportRefName.Name, loopStep));
            }
            else
            {
                // There are no more shortcuts now, we need to check at runtime that loopStep is negative for a descending loop and non-negative
                // for a non-descending loop
                guardClauseLines = guardClauseLines.Add(string.Format(
                    "(({0}.StrictLTE({1}, {2}) && {0}.StrictGTE({3}, 0))",
                    _supportRefName.Name,
                    loopStart,
                    loopEnd,
                    loopStep
                ));
                guardClauseLines = guardClauseLines.Add(string.Format(
                    "|| ({0}.StrictGT({1}, {2}) && {0}.StrictLT({3}, 0)))",
                    _supportRefName.Name,
                    loopStart,
                    loopEnd,
                    loopStep
                ));
            }

            // If there are non-constant constraint(s) and error-trapping is enabled and thee constraint evaluation fails, then VBScript will
            // enter the loop once (only once, and it will not initialise the loop variable when doing so). In this case, we need to bypass
            // the guard clause - it won't make sense anyway to compare loop constaints if we know that the constraint evaluation failed!
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                // If the guard clases have already been split up over multiple lines then insert a fresh line, otherwise just bang it in
                // front of the single line (unless there is no content, in which case this will become the only guard clause line)
                if (!guardClauseLines.Any())
                    guardClauseLines = guardClauseLines.Add("(" + constraintsInitialisedFlagNameIfAny.Name + ")");
                else if (guardClauseLines.Count == 1)
                {
                    guardClauseLines = new NonNullImmutableList<string>(new[]
                    {
                        string.Format(
                            "(!{0} || {1})",
                            constraintsInitialisedFlagNameIfAny.Name,
                            guardClauseLines.Single()
                        )
                    });
                }
                else
                {
                    guardClauseLines = new NonNullImmutableList<string>(
                        new[] { "(!" + constraintsInitialisedFlagNameIfAny.Name + " ||" }
                        .Concat(guardClauseLines.Take(guardClauseLines.Count - 1))
                        .Concat(new[] { guardClauseLines.Last() + ")" })
                    );
                }
            }

            if (guardClauseLines.Any())
            {
                translationResult = translationResult.Add(new TranslatedStatement("if " + guardClauseLines.First(), indentationDepth));
                foreach (var guardClauseLine in guardClauseLines.Skip(1))
                    translationResult = translationResult.Add(new TranslatedStatement(guardClauseLine, indentationDepth));
                translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
                indentationDepth++;
            }
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(
                            "{0}.HANDLEERROR({1}, () => {{",
                            _supportRefName.Name,
                            scopeAccessInformation.ErrorRegistrationTokenIfAny.Name
                        ),
                        indentationDepth
                    ));
                indentationDepth++;
            }
            string continuationCondition;
            if (numericLoopStepValueIfAny != null)
            {
                if (numericLoopStepValueIfAny.Value >= 0)
                {
                    // Note: If loopEnd is a known numeric constant then we can render just its value (eg. "1") instead of its translated content, which
                    // would include type information (eg. "(Int16)1") since this will make the final output a little more succinct (and the type information
                    // that we're skipping will have no effect on the StrictLTE call). This is most obvious for really simple loops but it doesn't hurt to
                    // try to make the more complicated ones shorter (since they get increasingly complicated and verbose as dynamic loop constraints and
                    // potential error-trapping are added!).
                    continuationCondition = string.Format(
                        "{0}.StrictLTE({1}, {2})",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        (numericLoopEndValueIfAny == null) ? loopEnd : numericLoopEndValueIfAny.Value.ToString()
                    );
                }
                else
                {
                    // Note: If loopEnd is a known numeric constant then we can render just its value instead of its translated content - see note above
                    continuationCondition = string.Format(
                        "{0}.StrictGTE({1}, {2})",
                        _supportRefName.Name,
                        rewrittenLoopVariableName,
                        (numericLoopEndValueIfAny == null) ? loopEnd : numericLoopEndValueIfAny.Value.ToString()
                    );
                }
            }
            else
            {
                // Note: If loopEnd is a known numeric constant then we can render just its value instead of its translated content - see note above
                continuationCondition = string.Format(
                    "({0}.StrictGTE({3}, 0) && {0}.StrictLTE({1}, {2})) || ({0}.StrictLT({3}, 0) && {0}.StrictGTE({1}, {2}))",
                    _supportRefName.Name,
                    rewrittenLoopVariableName,
                    (numericLoopEndValueIfAny == null) ? loopEnd : numericLoopEndValueIfAny.Value.ToString(),
                    loopStep
                );
            }
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                // If there is a chance that the constraint evaluation will fail but that processing may continue (ie. there's an ON ERROR RESUME NEXT
                // in the current scope that may affect things) then the continuation condition gets a special condition that it will not be checked
                // if the constraint evaluation did indeed fail. In such a case, VBScript will enter the loop once (without altering the loop variable
                // value). To replicate this behaviour, if constraint evaluation fails then the loop WILL be entered once without the loop variable
                // being altered and without the continuation condition being considered - instead, there is a "break" at the end of the loop which
                // will be executed if the constraint evaluation failed, ensuring that the loop was indeed processed once, consistent with VBScript.
                continuationCondition = string.Format(
                    "!{0} || ({1})",
                    constraintsInitialisedFlagNameIfAny.Name,
                    continuationCondition
                );
            }
            string loopIncrementWithLeadingSpaceIfNonBlank;
            if ((numericLoopStepValueIfAny != null) && (numericLoopStepValueIfAny.Value == 0))
                loopIncrementWithLeadingSpaceIfNonBlank = "";
            else
            {
                if ((numericLoopStepValueIfAny != null) && (numericLoopStepValueIfAny.Value < 0))
                {
                    // If the step is known to be a negative numeric constant value, then use SUBT instead of ADD. If it's a small step value then it
                    // will cast to an Int16, so using SUBT instead of ADD allows the following (slight) improvement to readability:
                    //   _env.i = _.ADD(_env.i, (Int16)(-1))
                    //   _env.i = _.SUBT(_env.i, (Int16)1)
                    loopIncrementWithLeadingSpaceIfNonBlank = string.Format(
                        " {0} = {2}.SUBT({0}, {1})",
                        rewrittenLoopVariableName,
                        numericLoopStepValueIfAny.GetNegative().AsCSharpValue(),
                        _supportRefName.Name
                    );
                }
                else
                {
                    // Another shortcut we can take is if the step value is known to be a non-negative numeric constant, we can just render its value
                    // out, rather than its full translated content (eg. "1" compared to "(Int16)1", since its type will not have any effect on the
                    // ADD action and rendering its numeric value only is more succinct). CORRECTION: This is not true, since "CInt(1) + CLng(1)"
                    // will return a value of type "Long" (whereas "CInt(1) + CInt(1)" will return a value of type "Integer") - so this type
                    // information IS important. I'm leaving this entire comment (the wrong assumption and the correction) so that there's
                    // no chance of me coming back in the future and thinking I can change it back again!
                    loopIncrementWithLeadingSpaceIfNonBlank = string.Format(
                        " {0} = {2}.ADD({0}, {1})",
                        rewrittenLoopVariableName,
                        loopStep,
                        _supportRefName.Name
                    );
                }
            }
            string loopVarInitialiser;
            if (constraintsInitialisedFlagNameIfAny == null)
            {
                loopVarInitialiser = string.Format(
                    "{0} = {1}",
                    rewrittenLoopVariableName,
                    loopStart
                );
            }
            else
            {
                // If error-trapping is enabled and loop constaint evaluation fails, then the loop must be entered once but the loop variable's value may
                // not be altered. This condition deals with that case; the loop variable is only set to the start value if all of the constraints were
                // successfully initialised. Note that constraintsInitialisedFlagNameIfAny will only be non-null if error-trapping may take affect and
                // if the loop constraints are not numeric constants known at translation time.
                loopVarInitialiser = string.Format(
                    "{0} = {1} ? {2} : {0}",
                    rewrittenLoopVariableName,
                    constraintsInitialisedFlagNameIfAny.Name,
                    loopStart
                );
            }
            if ((constraintsInitialisedFlagNameIfAny == null) && (numericLoopStepValueIfAny != null))
            {
                // If there is no complicated were-loop-constraints-successfully-initialised logic to worry about (meaning that either error-trapping was
                // not enabled or that all of the constraints were numeric constants), then render a nice and simple single-line loop construct.
                // Note: There is no space before {2} so that if there is no loop increment required then the output doesn't look like it's missing
                // something (this may be the case if the loop step is zero)
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(
                        "for ({0}; {1};{2})",
                        loopVarInitialiser,
                        continuationCondition,
                        loopIncrementWithLeadingSpaceIfNonBlank
                    ),
                    indentationDepth
                ));
            }
            else
            {
                // If there IS some possibility that loop constraint evaluation may fail at runtime, then the loop construct becomes much more complicated
                // and so benefits from being broken over multiple lines. There is still a chance that the loop increment could be blank, so consider that
                // and either break over two or three lines.
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(
                        "for ({0};",
                        loopVarInitialiser
                    ),
                    indentationDepth
                ));
                translationResult = translationResult.Add(new TranslatedStatement(
                    continuationCondition + ";" + ((loopIncrementWithLeadingSpaceIfNonBlank == "") ? ")" : ""),
                    indentationDepth + 1
                ));
                if (loopIncrementWithLeadingSpaceIfNonBlank != "")
                {
                    translationResult = translationResult.Add(new TranslatedStatement(
                        loopIncrementWithLeadingSpaceIfNonBlank.Trim() + ")",
                        indentationDepth + 1
                    ));
                }
            }
            translationResult = translationResult.Add(new TranslatedStatement("{", indentationDepth));
            var earlyExitNameIfAny = GetEarlyExitNameIfRequired(forBlock, scopeAccessInformation);
            if (earlyExitNameIfAny != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format("var {0} = false;", earlyExitNameIfAny.Name),
                    indentationDepth + 1
                ));
            }
            translationResult = translationResult.Add(
                Translate(
                    forBlock.Statements.ToNonNullImmutableList(),
                    scopeAccessInformation.SetParent(forBlock),
                    earlyExitNameIfAny,
                    indentationDepth + 1
                )
            );
            if (constraintsInitialisedFlagNameIfAny != null)
            {
                // If error-trapping is enabled and loop constaint evaluation failed, then the loop is entered once but the loop variable value is not
                // altered and continuation criteria are never checked - instead, the loop exits after one pass in these circumstances
                translationResult = translationResult
                    .Add(new TranslatedStatement("if (!" + constraintsInitialisedFlagNameIfAny.Name + ")", indentationDepth + 1))
                    .Add(new TranslatedStatement("break;", indentationDepth + 2));
            }
            translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny != null)
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("});", indentationDepth));
            }
            if (guardClauseLines.Any())
            {
                indentationDepth--;
                translationResult = translationResult.Add(new TranslatedStatement("}", indentationDepth));
            }
            foreach (var undeclaredVariable in undeclaredVariableReferencesAccessedByLoopConstraints)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
            var earlyExitFlagNamesToCheck = scopeAccessInformation.StructureExitPoints
                .Where(e => e.ExitEarlyBooleanNameIfAny != null)
                .Select(e => e.ExitEarlyBooleanNameIfAny.Name);
            if (earlyExitFlagNamesToCheck.Any())
            {
                // Perform early-exit checks for any scopeAccessInformation.StructureExitPoints - if this is FOR loop inside a DO..LOOP loop and an
                // EXIT DO was encountered within the FOR that must refer to the containing DO, then the FOR loop will have been broken out of, but
                // also a flag set that means that we must break further to get out of the DO loop.
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
            return translationResult.AddUndeclaredVariables(undeclaredVariableReferencesAccessedByLoopConstraints);
		}

        /// <summary>
        /// If the expression is guaranteed to return a true numeric value (not null, not empty, not a boolean) then we don't need to wrap it in a NUM call
        /// (if the expression does not meet these criteria then this function will return content that does include a NUM call)
        /// </summary>
        private TranslatedStatementContentDetails WrapInNUMCallIfRequired(Expression expression, ScopeAccessInformation scopeAccessInformation)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            var translatedExpressionContent = _statementTranslator.Translate(
                expression,
                scopeAccessInformation,
                ExpressionReturnTypeOptions.NotSpecified,
                _logger.Warning
            );
            if (IsCallingBuiltInNumberReturningFunction(expression, scopeAccessInformation))
                return translatedExpressionContent;
            return new TranslatedStatementContentDetails(
                string.Format(
                    "{0}.NUM({1})",
                    _supportRefName.Name,
                    translatedExpressionContent.TranslatedContent
                ),
                translatedExpressionContent.VariablesAccessed
            );
        }

        private bool IsCallingBuiltInNumberReturningFunction(Expression expression, ScopeAccessInformation scopeAccessInformation)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            var expressions =
                  StageTwoExpressionParsing.ExpressionGenerator.Generate(
                      expression.Tokens,
                      (scopeAccessInformation.DirectedWithReferenceIfAny == null) ? null : scopeAccessInformation.DirectedWithReferenceIfAny.AsToken(),
                      _logger.Warning
                  )
                  .ToArray();
            if ((expressions.Length != 1) || (expressions[0].Segments.Count() != 1))
                return false;
            var callExpression = expressions[0].Segments.Single() as StageTwoExpressionParsing.CallExpressionSegment;
            if (callExpression == null)
                return false;
            if (callExpression.MemberAccessTokens.Count() != 1)
                return false;
            var builtInFunctionToken = callExpression.MemberAccessTokens.Single() as BuiltInFunctionToken;
            return (builtInFunctionToken != null) && builtInFunctionToken.GuaranteedToReturnNumericContent;
        }

        private NumericValueToken TryToGetExpressionAsNumericConstant(Expression expression)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");

            var tokens = expression.Tokens.ToArray();
            if (tokens.Length != 1)
                return null;
            return tokens[0] as NumericValueToken;
        }

        private CSharpName GetEarlyExitNameIfRequired(ForBlock forBlock, ScopeAccessInformation scopeAccessInformation)
        {
            if (forBlock == null)
                throw new ArgumentNullException("forBlock");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            if (!forBlock.ContainsLoopThatContainsMismatchedExitThatMustBeHandledAtThisLevel())
                return null;

            return _tempNameGenerator(new CSharpName("exitFor"), scopeAccessInformation);
        }

        private TranslationResult Translate(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, CSharpName earlyExitNameIfAny, int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // Add a StructureExitPoint entry for the current loop so that the "early-exit" logic described in the Translate method above is possible
            return base.TranslateCommon(
                base.GetWithinFunctionBlockTranslators(),
				blocks,
				scopeAccessInformation.AddStructureExitPoints(
                    earlyExitNameIfAny,
                    ScopeAccessInformation.ExitableNonScopeDefiningConstructOptions.For
                ),
				indentationDepth
			);
		}

        private class LoopConstraintInitialiser
        {
            public LoopConstraintInitialiser(CSharpName variableName, string initialisationContent)
            {
                if (variableName == null)
                    throw new ArgumentNullException("variableName");
                if (string.IsNullOrWhiteSpace(initialisationContent))
                    throw new ArgumentException("Null/blank initialisationContent specified");

                VariableName = variableName;
                InitialisationContent = initialisationContent;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public CSharpName VariableName { get; private set; }

            /// <summary>
            /// This will never be null or blank
            /// </summary>
            public string InitialisationContent { get; private set; }
        }

        private class LoopStartConstraintInitialiser : LoopConstraintInitialiser
        {
            public LoopStartConstraintInitialiser(CSharpName variableName, string initialisationContentIgnoringTypesOfOtherConstraints, string initialisationContent)
                : base(variableName, initialisationContent)
            {
                if (string.IsNullOrWhiteSpace(initialisationContentIgnoringTypesOfOtherConstraints))
                    throw new ArgumentException("Null/blank initialisationContentIgnoringTypesOfOtherConstraints specified");

                InitialisationContentIgnoringTypesOfOtherConstraints = initialisationContentIgnoringTypesOfOtherConstraints;
            }

            /// <summary>
            /// This will never be null or blank. It might be the same as InitialisationContent or it might be different, depending upon the source
            /// code being translated. If it is the same then there were no other loop constraints that could affect the loop start value. If it
            /// is different then there are other loop constraints which might affect it. It is important to be able to do the work in two parts
            /// when error-trapping is enabled (see where this class is used for more details).
            /// </summary>
            public string InitialisationContentIgnoringTypesOfOtherConstraints { get; private set; }
        }
    }
}
