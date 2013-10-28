using CSharpWriter.Lists;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class ScopeAccessInformation_Extensions
    {
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
			IHaveNestedContent parentIfAny,
			IDefineScope scopeDefiningParentIfAny,
            CSharpName parentReturnValueNameIfAny,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            blocks = blocks.FlattenAllAccessibleBlockLevelCodeBlocks();
            var variables = scopeInformation.Variables.AddRange(
                blocks
                    .Where(b => b is DimStatement) // This covers DIM, REDIM, PRIVATE and PUBLIC (they may all be considered the same for these purposes)
                    .Cast<DimStatement>()
                    .SelectMany(d => d.Variables.Select(v => v.Name))
            );
            if (scopeDefiningParentIfAny != null)
                variables = variables.AddRange(scopeDefiningParentIfAny.ExplicitScopeAdditions);

            return new ScopeAccessInformation(
				parentIfAny,
				scopeDefiningParentIfAny,
                parentReturnValueNameIfAny,
                scopeInformation.Classes.AddRange(
                    blocks
                        .Where(b => b is ClassBlock)
                        .Cast<ClassBlock>()
                        .Select(c => c.Name)
                ),
                scopeInformation.Functions.AddRange(
                    blocks
                        .Where(b => (b is FunctionBlock) || (b is SubBlock))
                        .Cast<AbstractFunctionBlock>()
                        .Select(f => f.Name)
                ),
                scopeInformation.Properties.AddRange(
                    blocks
                        .Where(b => b is PropertyBlock)
                        .Cast<PropertyBlock>()
                        .Select(p => p.Name)
                ),
                variables
            );
        }

        /// <summary>
        /// If the parentIfAny is scope-defining then both the parentIfAny and scopeDefiningParentIfAny references will be set to it, this is a convenience
        /// method to save having to specify it explicitly for both
        /// </summary>
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IDefineScope parentIfAny,
            CSharpName parentReturnValueNameIfAny,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            return Extend(scopeInformation, parentIfAny, parentIfAny, parentReturnValueNameIfAny, blocks);
        }

        /// <summary>
        /// If the parentIfAny is scope-defining then both the parentIfAny and scopeDefiningParentIfAny references will be set to it, this is a convenience
        /// method to save having to specify it explicitly for both (for cases where the parent scope - if any - does not have a return value)
        /// </summary>
        public static ScopeAccessInformation Extend(this ScopeAccessInformation scopeInformation, IDefineScope parentIfAny, NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            return Extend(scopeInformation, parentIfAny, null, blocks);
        }

        /// <summary>
        /// Filter a set of NameTokens to only those that have not been declared in the specified scope
        /// </summary>
        public static NonNullImmutableList<NameToken> GetUndeclaredVariables(this ScopeAccessInformation scopeInformation, NonNullImmutableList<NameToken> variablesAccessed)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (variablesAccessed == null)
                throw new ArgumentNullException("memberCallVariablesAccessed");

            // TODO: If the scopeDefiningParentIfAny is a FUNCTION or PROPERTY then it will have a return value variable defined which needs to be included in
            // this declaredReferences set
            var declaredReferences =
                scopeInformation.Classes
                .AddRange(scopeInformation.Functions)
                .AddRange(scopeInformation.Properties)
                .AddRange(scopeInformation.Variables);
            return variablesAccessed
                .Where(v => !declaredReferences.Any(r => r.Content.Equals(v.Content, StringComparison.InvariantCultureIgnoreCase)))
                .ToNonNullImmutableList();
        }
    }
}
