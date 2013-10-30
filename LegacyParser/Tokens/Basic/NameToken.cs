using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This class is used by FunctionBlocks, ValueSettingStatements and other places where content is required that must represent a VBScript name
    /// </summary>
    [Serializable]
    public class NameToken : AtomToken
    {
        public NameToken(string content, int lineIndex) : this(content, WhiteSpaceBehaviourOptions.Disallow, lineIndex) { }

        protected NameToken(string content, WhiteSpaceBehaviourOptions whiteSpaceBehaviour, int lineIndex) : base(content, whiteSpaceBehaviour, lineIndex)
        {
            // If this constructor is being called from a type derived from NameToken (eg. EscapedNameToken) then assume that all validation has been
            // performed in its constructor. If this constructor is being called to instantiate a new NameToken (and NOT a class derived from it) then
            // use the AtomToken's TryToGetAsRecognisedType method to try to ensure that this content is valid as a name and should not be for a token
            // of another type. This is process is kind of hokey but I'm trying to layer on a little additional type safety to some very old code so
            // I'm willing to live with this approach to it (the base class - the AtomToken - having knowledge of all of the derived types is not a
            // great design decision).
            if (this.GetType() == typeof(NameToken))
            {
                var recognisedType = TryToGetAsRecognisedType(content, lineIndex);
                if (recognisedType != null)
                    throw new ArgumentException("Invalid content for a NameToken");
            }
        }
    }
}
