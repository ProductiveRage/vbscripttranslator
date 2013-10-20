using CSharpWriter.Lists;
using System;
using System.Collections.Generic;

namespace CSharpWriter
{
    public static class TranslationResult_Extensions
    {
        public static TranslationResult Add(this TranslationResult source, TranslatedStatement toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements.Add(toAdd),
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult Add(this TranslationResult source, TranslationResult toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements.AddRange(toAdd.TranslatedStatements),
                source.ExplicitVariableDeclarations.AddRange(toAdd.ExplicitVariableDeclarations),
                source.UndeclaredVariablesAccessed.AddRange(toAdd.UndeclaredVariablesAccessed)
            );
        }

        public static TranslationResult Add(this TranslationResult source, IEnumerable<VariableDeclaration> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements,
                source.ExplicitVariableDeclarations.AddRange(toAdd.ToNonNullImmutableList()),
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult Add(this TranslationResult source, TranslatedStatement toAddBefore, TranslationResult toAdd, TranslatedStatement toAddAfter)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAddBefore == null)
                throw new ArgumentNullException("toAddBefore");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");
            if (toAddAfter == null)
                throw new ArgumentNullException("toAddAfter");

            return new TranslationResult(
                source.TranslatedStatements
                    .Add(toAddBefore)
                    .AddRange(toAdd.TranslatedStatements)
                    .Add(toAddAfter),
                source.ExplicitVariableDeclarations.AddRange(toAdd.ExplicitVariableDeclarations),
                source.UndeclaredVariablesAccessed.AddRange(toAdd.UndeclaredVariablesAccessed)
            );
        }
    }
}
