using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter
{
    public class TranslationResult
    {
        public TranslationResult(
            NonNullImmutableList<TranslatedStatement> translatedStatements,
            NonNullImmutableList<VariableDeclaration> explicitVariableDeclarations,
            NonNullImmutableList<NameToken> undeclaredVariablesAccessed)
        {
            if (translatedStatements == null)
                throw new ArgumentNullException("translatedStatements");
            if (explicitVariableDeclarations == null)
                throw new ArgumentNullException("explicitVariableDeclarations");
            if (undeclaredVariablesAccessed == null)
                throw new ArgumentNullException("undeclaredVariablesAccessed");

            TranslatedStatements = translatedStatements;
            ExplicitVariableDeclarations = explicitVariableDeclarations;
            UndeclaredVariablesAccessed = undeclaredVariablesAccessed;
        }

        public static TranslationResult Empty
        {
            get
            {
                return new TranslationResult(
                    new NonNullImmutableList<TranslatedStatement>(),
                    new NonNullImmutableList<VariableDeclaration>(),
                    new NonNullImmutableList<NameToken>()
                );
            }
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<TranslatedStatement> TranslatedStatements { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<VariableDeclaration> ExplicitVariableDeclarations { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> UndeclaredVariablesAccessed { get; private set; }
    }
}
