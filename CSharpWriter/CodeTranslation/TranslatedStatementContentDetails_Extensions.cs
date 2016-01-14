using VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions;
using VBScriptTranslator.CSharpWriter.Lists;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation
{
    public static class TranslatedStatementContentDetails_Extensions
    {
        /// <summary>
        /// This will never be null
        /// </summary>
        public static NonNullImmutableList<NameToken> GetUndeclaredVariablesAccessed(
            this TranslatedStatementContentDetails source,
            ScopeAccessInformation scopeAccessInformation,
            VBScriptNameRewriter nameRewriter)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            return source.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(v, nameRewriter))
                .ToNonNullImmutableList();
        }
    }
}
