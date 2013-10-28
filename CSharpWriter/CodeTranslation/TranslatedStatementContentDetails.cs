using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class TranslatedStatementContentDetails
    {
        public TranslatedStatementContentDetails(string translatedContent, NonNullImmutableList<NameToken> variablesAccesed)
        {
            if (string.IsNullOrWhiteSpace(translatedContent))
                throw new ArgumentException("Null/blank translatedContent specified");
            if (variablesAccesed == null)
                throw new ArgumentNullException("variablesAccesed");

            TranslatedContent = translatedContent;
            VariablesAccesed = variablesAccesed;
        }

        /// <summary>
        /// This will never return null or blank
        /// </summary>
        public string TranslatedContent { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> VariablesAccesed { get; private set; }
    }
}
