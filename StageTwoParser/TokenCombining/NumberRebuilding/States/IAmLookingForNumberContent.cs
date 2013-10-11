using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public interface IAmLookingForNumberContent
    {
        TokenProcessResult Process(IToken token, PartialNumberContent numberContent);
    }
}
