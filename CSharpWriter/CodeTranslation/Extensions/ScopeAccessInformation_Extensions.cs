using CSharpWriter.Lists;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
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

        public static ScopeAccessInformation SetErrorRegistrationToken(this ScopeAccessInformation scopeAccessInformation, CSharpName errorRegistrationTokenIfAny)
        {
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            return new ScopeAccessInformation(
                scopeAccessInformation.ParentIfAny,
                scopeAccessInformation.ScopeDefiningParentIfAny,
                scopeAccessInformation.ParentReturnValueNameIfAny,
                errorRegistrationTokenIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions,
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables
            );
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
        /// TODO
        /// </summary>
        public static CSharpName GetNameOfTargetContainerIfAnyRequired(
            this ScopeAccessInformation scopeAccessInformation,
            string rewrittenTargetName,
            CSharpName envRefName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter)
        {
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (string.IsNullOrWhiteSpace(rewrittenTargetName))
                throw new ArgumentException("Null/blank rewrittenTargetName specified");
            if (envRefName == null)
                throw new ArgumentNullException("envRefName");
            if (outerRefName == null)
                throw new ArgumentNullException("outerRefName");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            var targetReferenceDetailsIfAvailable = scopeAccessInformation.TryToGetDeclaredReferenceDetails(rewrittenTargetName, nameRewriter);
            if (targetReferenceDetailsIfAvailable == null)
            {
                if (scopeAccessInformation.ScopeLocation == ScopeLocationOptions.WithinFunctionOrProperty)
                {
                    // If an undeclared variable is accessed within a function (or property) then it is treated as if it was declared to be restricted
                    // to the current scope, so the nameOfTargetContainerIfRequired should be null in this case (this means that the UndeclaredVariables
                    // data returned from this process should be translated into locally-scoped DIM statements at the top of the function / property).
                    return null;
                }
                return envRefName;
            }
            else if (targetReferenceDetailsIfAvailable.ReferenceType == ReferenceTypeOptions.ExternalDependency)
                return envRefName;
            else if (targetReferenceDetailsIfAvailable.ScopeLocation == ScopeLocationOptions.OutermostScope)
            {
                // 2014-01-06 DWR: Used to only apply this logic if the target reference was in the OutermostScope and we were currently inside a
                // class but I'm restructuring the outer scope so that declared variables and functions are inside a class that the outermost scope
                // references in an identical manner to the class functions (and properties) so the outerRefName should used every time that an
                // OutermostScope reference is accessed
                return outerRefName;
            }
            return null;
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
                throw new ArgumentException("Null/blank rewrittenTargetName specified");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            if (scopeInformation.ScopeDefiningParentIfAny != null)
            {
                if (scopeInformation.ScopeDefiningParentIfAny.ExplicitScopeAdditions.Any(t => nameRewriter.GetMemberAccessTokenName(t) == rewrittenTargetName))
                {
                    // ExplicitScopeAdditions should be things such as function arguments, so they will share the same ScopeLocation as the
                    // current scopeInformation reference
                    return new DeclaredReferenceDetails(ReferenceTypeOptions.Variable, scopeInformation.ScopeLocation);
                }
            }

            var firstExternalDependencyMatch = scopeInformation.ExternalDependencies
                .FirstOrDefault(t => nameRewriter.GetMemberAccessTokenName(t) == rewrittenTargetName);
            if (firstExternalDependencyMatch != null)
                return new DeclaredReferenceDetails(ReferenceTypeOptions.ExternalDependency, ScopeLocationOptions.OutermostScope);

            var scopedNameTokens =
                scopeInformation.Classes.Select(t => Tuple.Create(t, ReferenceTypeOptions.Class))
                .Concat(scopeInformation.Functions.Select(t => Tuple.Create(t, ReferenceTypeOptions.Function)))
                .Concat(scopeInformation.Properties.Select(t => Tuple.Create(t, ReferenceTypeOptions.Property)))
                .Concat(scopeInformation.Variables.Select(t => Tuple.Create(t, ReferenceTypeOptions.Variable)));

            // There could be references matching the requested name in multiple scopes, start from the closest and work outwards
            var possibleScopes = new[]
            {
                ScopeLocationOptions.WithinFunctionOrProperty,
                ScopeLocationOptions.WithinClass,
                ScopeLocationOptions.OutermostScope
            };
            foreach (var scope in possibleScopes)
            {
                var firstMatch = scopedNameTokens
                    .Where(t => t.Item1.ScopeLocation == scope)
                    .FirstOrDefault(t => nameRewriter.GetMemberAccessTokenName(t.Item1) == rewrittenTargetName);
                if (firstMatch != null)
                    return new DeclaredReferenceDetails(firstMatch.Item2, firstMatch.Item1.ScopeLocation);
            }
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
        }
    }
}
