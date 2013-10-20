using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    /// <summary>
    /// This must be responsible for translating any NameToken into a string that is legal for use as a C# identifier. It not expect a null name reference
    /// and must never return a null value. It is responsible for returning consistent names regardless of the case of the input value, to deal with the
    /// fact that C#  is case-sensitive and VBScript is not.
    /// </summary>
    public delegate CSharpName VBScriptNameRewriter(NameToken name);

    /// <summary>
    /// During translation, temporary variables may be required. This delegate is responsible for returning names that are guaranteed to be unique. The
    /// mechanism for implementing this must work with the VBScriptNameRewriter mechanism since there must be no overlap in the returned values. If an
    /// optionalPrefix value is specified then the returned name must begin with this (if null is specified then it must be ignored).
    /// </summary>
    public delegate CSharpName TempValueNameGenerator(CSharpName optionalPrefix);
}
