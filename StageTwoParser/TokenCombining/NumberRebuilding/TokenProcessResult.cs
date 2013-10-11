using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding
{
    public class TokenProcessResult
    {
        public TokenProcessResult(PartialNumberContent numberContent, IEnumerable<IToken> processedTokens, IAmLookingForNumberContent nextProcessor)
        {
            if (numberContent == null)
                throw new ArgumentNullException("numberContent");
            if (processedTokens == null)
                throw new ArgumentNullException("processedTokens");
            if (nextProcessor == null)
                throw new ArgumentNullException("nextProcessor");

            NumberContent = numberContent;
            ProcessedTokens = processedTokens.ToList().AsReadOnly();
            if (ProcessedTokens.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in processedTokens set");
            NextProcessor = nextProcessor;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public PartialNumberContent NumberContent { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references
        /// </summary>
        public IEnumerable<IToken> ProcessedTokens { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public IAmLookingForNumberContent NextProcessor { get; private set; }
    }
}
