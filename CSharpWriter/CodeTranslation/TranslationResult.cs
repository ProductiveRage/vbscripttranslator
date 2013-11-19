using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class TranslationResult
    {
        public TranslationResult(
            NonNullImmutableList<TranslatedStatement> translatedStatements,
            NonNullImmutableList<VariableDeclaration> explicitVariableDeclarations,
            NonNullImmutableList<NameToken> environmentVariablesAccessed)
        {
            if (translatedStatements == null)
                throw new ArgumentNullException("translatedStatements");
            if (explicitVariableDeclarations == null)
                throw new ArgumentNullException("explicitVariableDeclarations");
            if (environmentVariablesAccessed == null)
                throw new ArgumentNullException("environmentVariablesAccessed");

            TranslatedStatements = translatedStatements;
            ExplicitVariableDeclarations = explicitVariableDeclarations;
            EnvironmentVariablesAccessed = environmentVariablesAccessed;
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
        /// This will never be null (this is a set of all variables that are referenced but never declared, in order to deal with this in the translated
        /// C# code an environment reference is maintained and passed between classes which makes these references available as properties)
        /// </summary>
        public NonNullImmutableList<NameToken> EnvironmentVariablesAccessed { get; private set; }
    }
}
