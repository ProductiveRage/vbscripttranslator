using CSharpWriter.Lists;
using System;
using System.Collections.Generic;
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
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeLocationOptions blocksScopeLocation)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), blocksScopeLocation))
                throw new ArgumentOutOfRangeException("blocksScopeLocation");

            blocks = FlattenAllAccessibleBlockLevelCodeBlocks(blocks);
            var variables = scopeInformation.Variables.AddRange(
                blocks
                    .Where(b => b is DimStatement) // This covers DIM, REDIM, PRIVATE and PUBLIC (they may all be considered the same for these purposes)
                    .Cast<DimStatement>()
                    .SelectMany(d => d.Variables.Select(v => new ScopedNameToken(v.Name.Content, v.Name.LineIndex, blocksScopeLocation)))
            );
            if (scopeDefiningParentIfAny != null)
            {
                variables = variables.AddRange(
                    scopeDefiningParentIfAny.ExplicitScopeAdditions.Select(v => new ScopedNameToken(v.Content, v.LineIndex, blocksScopeLocation))
                );
            }

            return new ScopeAccessInformation(
				parentIfAny,
				scopeDefiningParentIfAny,
                parentReturnValueNameIfAny,
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
                variables,
                blocksScopeLocation
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
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeLocationOptions blocksScopeLocation)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), blocksScopeLocation))
                throw new ArgumentOutOfRangeException("blocksScopeLocation");

            return Extend(scopeInformation, parentIfAny, parentIfAny, parentReturnValueNameIfAny, blocks, blocksScopeLocation);
        }

        /// <summary>
        /// If the parentIfAny is scope-defining then both the parentIfAny and scopeDefiningParentIfAny references will be set to it, this is a convenience
        /// method to save having to specify it explicitly for both (for cases where the parent scope - if any - does not have a return value)
        /// </summary>
        public static ScopeAccessInformation Extend(
            this ScopeAccessInformation scopeInformation,
            IDefineScope parentIfAny,
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeLocationOptions blocksScopeLocation)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (blocks == null)
                throw new ArgumentNullException("blocks");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), blocksScopeLocation))
                throw new ArgumentOutOfRangeException("blocksScopeLocation");

            return Extend(scopeInformation, parentIfAny, null, blocks, blocksScopeLocation);
        }

        public static bool IsDeclaredReference(this ScopeAccessInformation scopeInformation, string rewrittenTargetName, VBScriptNameRewriter nameRewriter)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (string.IsNullOrWhiteSpace(rewrittenTargetName))
                throw new ArgumentException("Null/blank rewrittenTargetName specified");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            return TryToGetDeclaredReferenceDetails(scopeInformation, rewrittenTargetName, nameRewriter) != null;
        }

        /// <summary>
        /// Try to retrieve information about a name token (that has been passed through the specified nameRewriter). If there is nothing matching it in the
        /// current scope then null will be returned.
        /// </summary>
        public static DeclaredReferenceDetails TryToGetDeclaredReferenceDetails(
            this ScopeAccessInformation scopeInformation,
            string rewrittenTargetName,
            VBScriptNameRewriter nameRewriter)
        {
            if (scopeInformation == null)
                throw new ArgumentNullException("scopeInformation");
            if (string.IsNullOrWhiteSpace(rewrittenTargetName))
                throw new ArgumentException("Null/blank targetName specified");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            if (scopeInformation.ScopeDefiningParentIfAny != null)
            {
                if (scopeInformation.ScopeDefiningParentIfAny.ExplicitScopeAdditions.Any(t => nameRewriter.GetMemberAccessTokenName(t) == rewrittenTargetName))
                {
                    // ExplicitScopeAdditions should be things such as function arguments, so they will share the same ScopeLocation as the
                    // current scopeInformation reference
                    return new DeclaredReferenceDetails(DeclaredReferenceDetails.ReferenceTypeOptions.Variable, scopeInformation.ScopeLocation);
                }
            }

            var scopedNameTokens =
                scopeInformation.Classes.Select(t => Tuple.Create(t, DeclaredReferenceDetails.ReferenceTypeOptions.Class))
                .Concat(scopeInformation.Functions.Select(t => Tuple.Create(t, DeclaredReferenceDetails.ReferenceTypeOptions.Function)))
                .Concat(scopeInformation.Properties.Select(t => Tuple.Create(t, DeclaredReferenceDetails.ReferenceTypeOptions.Property)))
                .Concat(scopeInformation.Variables.Select(t => Tuple.Create(t, DeclaredReferenceDetails.ReferenceTypeOptions.Variable)));
            
            var firstClassScopedMatch = scopedNameTokens
                .Where(t => t.Item1.ScopeLocation == ScopeLocationOptions.WithinClass)
                .FirstOrDefault(t => nameRewriter.GetMemberAccessTokenName(t.Item1) == rewrittenTargetName);
            if (firstClassScopedMatch != null)
                return new DeclaredReferenceDetails(firstClassScopedMatch.Item2, firstClassScopedMatch.Item1.ScopeLocation);
            
            var firstOutermostScopedMatch = scopedNameTokens
                .Where(t => t.Item1.ScopeLocation == ScopeLocationOptions.OutermostScope)
                .FirstOrDefault(t => nameRewriter.GetMemberAccessTokenName(t.Item1) == rewrittenTargetName);
            if (firstOutermostScopedMatch != null)
                return new DeclaredReferenceDetails(firstOutermostScopedMatch.Item2, firstOutermostScopedMatch.Item1.ScopeLocation);
            
            return null;
        }

        public class DeclaredReferenceDetails
        {
            public DeclaredReferenceDetails(ReferenceTypeOptions referenceType, ScopeLocationOptions scopeLocation)
            {
                if (!Enum.IsDefined(typeof(ReferenceTypeOptions), referenceType))
                    throw new ArgumentOutOfRangeException("referenceType");
                if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                    throw new ArgumentOutOfRangeException("scopeLocation");

                ReferenceType = referenceType;
                ScopeLocation = scopeLocation;
            }

            public ReferenceTypeOptions ReferenceType { get; private set; }
            public ScopeLocationOptions ScopeLocation { get; private set; }

            public enum ReferenceTypeOptions
            {
                Class,
                Function,
                Property,
                Variable
            }
        }
    }
}
