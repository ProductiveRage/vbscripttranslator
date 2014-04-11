using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public interface IAmLookingForNumberContent
    {
        TokenProcessResult Process(IEnumerable<IToken> tokens, PartialNumberContent numberContent);
    }
}
