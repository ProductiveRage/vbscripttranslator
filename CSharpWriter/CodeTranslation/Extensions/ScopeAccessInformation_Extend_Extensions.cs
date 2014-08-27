using CSharpWriter.Lists;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class ScopeAccessInformation_Extend_Extensions
    {
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IHaveNestedContent parentIfAny,
            IDefineScope scopeDefiningParentIfAny,
            CSharpName parentReturnValueNameIfAny,
            CSharpName errorRegistrationTokenIfAny,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            var blocksScopeLocation = (scopeDefiningParentIfAny == null) ? scopeInformation.ScopeLocation : scopeDefiningParentIfAny.Scope;
            blocks = FlattenAllAccessibleBlockLevelCodeBlocks(blocks);
            var variables = scopeInformation.Variables.AddRange(
                blocks
                    .Where(b => b is DimStatement) // This covers DIM, REDIM, PRIVATE and PUBLIC (they may all be considered the same for these purposes)
                    .Cast<DimStatement>()
                    .SelectMany(d => d.Variables.Select(v => new ScopedNameToken(
                        v.Name.Content,
                        v.Name.LineIndex,
                        blocksScopeLocation
                    )))
            );
            if (scopeDefiningParentIfAny != null)
            {
                variables = variables.AddRange(
                    scopeDefiningParentIfAny.ExplicitScopeAdditions
                        .Select(v => new ScopedNameToken(
                            v.Content,
                            v.LineIndex,
                            blocksScopeLocation
                        )
                    )
                );
            }

            return new ScopeAccessInformation(
                parentIfAny,
                scopeDefiningParentIfAny,
                parentReturnValueNameIfAny,
                errorRegistrationTokenIfAny,
                scopeInformation.ExternalDependencies,
                scopeInformation.Classes.AddRange(
                    blocks
                        .Where(b => b is ClassBlock)
                        .Cast<ClassBlock>()
                        .Select(c => new ScopedNameToken(c.Name.Content, c.Name.LineIndex, ScopeLocationOptions.OutermostScope)) // These are always OutermostScope
                ),
                scopeInformation.Functions.AddRange(
                    blocks
                        .Where(b => (b is FunctionBlock) || (b is SubBlock))
                        .Cast<AbstractFunctionBlock>()
                        .Select(b => new ScopedNameToken(b.Name.Content, b.Name.LineIndex, blocksScopeLocation))
                ),
                scopeInformation.Properties.AddRange(
                    blocks
                        .Where(b => b is PropertyBlock)
                        .Cast<PropertyBlock>()
                        .Select(p => new ScopedNameToken(p.Name.Content, p.Name.LineIndex, ScopeLocationOptions.WithinClass)) // These are always WithinClass
                ),
                variables
            );
        }

        private static NonNullImmutableList<ICodeBlock> FlattenAllAccessibleBlockLevelCodeBlocks(NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            var flattenedBlocks = new NonNullImmutableList<ICodeBlock>();
            foreach (var block in blocks)
            {
                flattenedBlocks = flattenedBlocks.Add(block);

                var parentBlock = block as IHaveNestedContent;
                if (parentBlock == null)
                    continue;

                if (parentBlock is IDefineScope)
                {
                    // If this defines scope then we can't expand the current scope by drilling into it - eg. if the current block
                    // is a class then it has nested statements but we can't access them directly (we can't call a function on a
                    // class without calling it on an instance of that class)
                    continue;
                }

                flattenedBlocks = flattenedBlocks.AddRange(
                    FlattenAllAccessibleBlockLevelCodeBlocks(
                        parentBlock.AllExecutableBlocks.ToNonNullImmutableList()
                    )
                );
            }
            return flattenedBlocks;
        }

        /// <summary>
        /// If the parentIfAny is scope-defining then both the parentIfAny and scopeDefiningParentIfAny references will be set to it, this is a convenience
        /// method to save having to specify it explicitly for both
        /// </summary>
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IDefineScope parentIfAny,
            CSharpName parentReturnValueNameIfAny,
            CSharpName errorRegistrationTokenIfAny,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            return Extend(scopeInformation, parentIfAny, parentIfAny, parentReturnValueNameIfAny, errorRegistrationTokenIfAny, blocks);
        }

        public static ScopeAccessInformation ExtendExternalDependencies(this ScopeAccessInformation scopeInformation, NonNullImmutableList<NameToken> externalDependencies)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");

            return new ScopeAccessInformation(
                scopeInformation.ParentIfAny,
                scopeInformation.ScopeDefiningParentIfAny,
                scopeInformation.ParentReturnValueNameIfAny,
                scopeInformation.ErrorRegistrationTokenIfAny,
                scopeInformation.ExternalDependencies.AddRange(externalDependencies),
                scopeInformation.Classes,
                scopeInformation.Functions,
                scopeInformation.Properties,
                scopeInformation.Variables
            );
        }

        public static ScopeAccessInformation ExtendVariables(this ScopeAccessInformation scopeInformation, NonNullImmutableList<ScopedNameToken> variables)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (variables == null)
                throw new ArgumentNullException("variables");

            return new ScopeAccessInformation(
                scopeInformation.ParentIfAny,
                scopeInformation.ScopeDefiningParentIfAny,
                scopeInformation.ParentReturnValueNameIfAny,
                scopeInformation.ErrorRegistrationTokenIfAny,
                scopeInformation.ExternalDependencies,
                scopeInformation.Classes,
                scopeInformation.Functions,
                scopeInformation.Properties,
                scopeInformation.Variables.AddRange(variables)
            );
        }

        /// <summary>
        /// If the parentIfAny is scope-defining then both the parentIfAny and scopeDefiningParentIfAny references will be set to it, this is a convenience
        /// method to save having to specify it explicitly for both (for cases where the parent scope - if any - does not have a return value)
        /// </summary>
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IDefineScope parentIfAny,
            NonNullImmutableList<ICodeBlock> blocks)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            return Extend(scopeInformation, parentIfAny, null, null, blocks);
        }
    }
}