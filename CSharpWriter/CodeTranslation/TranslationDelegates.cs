using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation
{
    /// <summary>
    /// This must be responsible for translating any NameToken into a string that is legal for use as a C# identifier. It should not expect a null name
    /// reference and must never return a null value. It is responsible for returning consistent names regardless of the case of the input value, to deal
    /// with the fact that C# is case-sensitive and VBScript is not.
    /// </summary>
    public delegate CSharpName VBScriptNameRewriter(NameToken name);

    /// <summary>
    /// During translation, temporary variables may be required. This delegate is responsible for returning names that are guaranteed to be unique. The
    /// mechanism for implementing this must work with the VBScriptNameRewriter mechanism since there must be no overlap in the returned values. If an
    /// optionalPrefix value is specified then the returned name must begin with this (if null is specified then it must be ignored). It is not expected
    /// that variables returned by this will be added to the ScopeAccessInformation, so they must be unique internally - meaning that no two calls to
    /// the method may return the same value - but can also ensure that the returned value do not clash with any variables that ARE in the data in
    /// the ScopeAccessInformation reference).
    /// </summary>
    public delegate CSharpName TempValueNameGenerator(CSharpName optionalPrefix, ScopeAccessInformation scopeAccessInformation);
}
