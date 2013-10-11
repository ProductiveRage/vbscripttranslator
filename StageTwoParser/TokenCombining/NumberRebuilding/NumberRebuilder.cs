using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States;
using VBScriptTranslator.StageTwoParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding
{
    /// <summary>
    /// Given the output of a StringBreaker/TokenBreaker pass, this will try to reconstruct numbers into single AtomTokens. Doing so will allow
    /// any remaining MemberAccessorOrDecimalPointToken to be replaced with MemberAccessorTokens (there will no longer be any ambiguity). As
    /// part of this processing, the "optional zeros" will be explicitly expressed (".1" will be replaced with "0.1"). All numeric values
    /// will be included in the output as NumericValueToken instances.
    /// </summary>
    public static class NumberRebuilder
    {
        public static IEnumerable<IToken> Rebuild(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            // At the beginning of a token set, if the first token is a minus sign then it is elligible to be part of a number. If it is not the
            // first token then it may only incorporated into a number if preceded by another operator or an opening brace (which effectively
            // would make the token the start of a new expression). The same principle applies to a decimal point. This is why the initial
            // processor is a PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber.
            var processor = (IAmLookingForNumberContent)PeriodOrMinusSignOrNumberCouldIndicateStartOfNumber.Instance;
            var numberContent = new PartialNumberContent();
            var rebuiltTokens = new List<IToken>();
            foreach (var token in tokens)
            {
                if (token == null)
                    throw new ArgumentException("Null reference encountered in tokens set");

                var result = processor.Process(token, numberContent);
                if (result.ProcessedTokens.Any())
                    rebuiltTokens.AddRange(result.ProcessedTokens);
                processor = result.NextProcessor;
                numberContent = result.NumberContent;
            }
            if (numberContent.Tokens.Any())
            {
                var numericValue = numberContent.TryToExpressNumberFromTokens();
                if (numericValue == null)
                    rebuiltTokens.AddRange(numberContent.Tokens);
                else
                    rebuiltTokens.Add(new NumericValueToken(numericValue.Value));
            }
            return rebuiltTokens.Select(t => (t is MemberAccessorOrDecimalPointToken) ? new MemberAccessorToken() : t);
        }
    }
}
