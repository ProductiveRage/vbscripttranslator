using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
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
                .Where(v => !scopeAccessInformation.IsDeclaredReference(nameRewriter.GetMemberAccessTokenName(v), nameRewriter))
                .ToNonNullImmutableList();
        }
    }
}
