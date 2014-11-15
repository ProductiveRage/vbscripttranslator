using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class ScopeAccessInformation_Reference_Extensions
    {
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
                if (scopeAccessInformation.ScopeDefiningParent.Scope == ScopeLocationOptions.WithinFunctionOrPropertyOrWith)
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

            // If the target corresponds to the containing "WITH" reference (if any) then use that ("WITH a: .Go: END WITH" is translated
            // approximately into "var w123 = a; w123.Go();" where the "w123" is the DirectedWithReferenceIfAny and so we don't need to
            // check for other variables or functions that may apply, it's the local variable WITH construct target.
            if (scopeInformation.DirectedWithReferenceIfAny != null)
            {
                // Note that WithinFunctionOrPropertyOrWith is always specified here for the scope location since the WITH target should
                // not be part of the "outer most scope" variable set like variables declared in that scope in the source script - this
                // target reference is not something that can be altered, it is set in the current scope and accessed directly.
                if (nameRewriter(scopeInformation.DirectedWithReferenceIfAny.AsToken()).Name == rewrittenTargetName)
                    return new DeclaredReferenceDetails(ReferenceTypeOptions.Variable, ScopeLocationOptions.WithinFunctionOrPropertyOrWith);
            }

            if (scopeInformation.ScopeDefiningParent != null)
            {
                if (scopeInformation.ScopeDefiningParent.ExplicitScopeAdditions.Any(t => nameRewriter.GetMemberAccessTokenName(t) == rewrittenTargetName))
                {
                    // ExplicitScopeAdditions should be things such as function arguments, so they will share the same ScopeLocation as the
                    // current scopeInformation reference
                    return new DeclaredReferenceDetails(ReferenceTypeOptions.Variable, scopeInformation.ScopeDefiningParent.Scope);
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
                ScopeLocationOptions.WithinFunctionOrPropertyOrWith,
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
    }
}
