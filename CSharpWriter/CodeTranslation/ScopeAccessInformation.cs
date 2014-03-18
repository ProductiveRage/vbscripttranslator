using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class ScopeAccessInformation
    {
        public ScopeAccessInformation(
			IHaveNestedContent parentIfAny,
			IDefineScope scopeDefiningParentIfAny,
            CSharpName parentReturnValueNameIfAny,
            NonNullImmutableList<NameToken> externalDependencies,
            NonNullImmutableList<ScopedNameToken> classes,
            NonNullImmutableList<ScopedNameToken> functions,
            NonNullImmutableList<ScopedNameToken> properties,
            NonNullImmutableList<ScopedNameToken> variables)
        {
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");
            if (classes == null)
                throw new ArgumentNullException("classes");
            if (functions == null)
                throw new ArgumentNullException("functions");
            if (properties == null)
                throw new ArgumentNullException("properties");
            if (variables == null)
                throw new ArgumentNullException("variables");

			if ((parentIfAny == null) && (scopeDefiningParentIfAny != null))
				throw new ArgumentException("If scopeDefiningParentIfAny is non-null then parentIfAny must be");

            ParentIfAny = parentIfAny;
			ScopeDefiningParentIfAny = scopeDefiningParentIfAny;
            ParentReturnValueNameIfAny = parentReturnValueNameIfAny;
            ExternalDependencies = externalDependencies;
            Classes = classes;
            Functions = functions;
            Properties = properties;
            Variables = variables;
        }

        public static ScopeAccessInformation Empty = new ScopeAccessInformation(
            null,
			null,
            null,
            new NonNullImmutableList<NameToken>(),
            new NonNullImmutableList<ScopedNameToken>(),
            new NonNullImmutableList<ScopedNameToken>(),
            new NonNullImmutableList<ScopedNameToken>(),
            new NonNullImmutableList<ScopedNameToken>()
        );

        /// <summary>
		/// /// This will be null if there is no scope-defining parent - ie. in the outermost scope
        /// </summary>
		public IHaveNestedContent ParentIfAny { get; private set; }

		/// <summary>
		/// This will be null if there is no scope-defining parent - eg. in the outermost scope, or within a non-scope-altering construct (such as an
		/// IF block) within that scope. This may be the same reference as ParentIfAny. If this is non-null then ParentIfAny will always be non-null,
		/// though it is possible for ParentIfAny to be non-null and this be null (eg. when inside an IF block in the outermost scope)
		/// </summary>
		public IDefineScope ScopeDefiningParentIfAny { get; private set; }

        /// <summary>
        /// This will be null if the ScopeDefiningParentIfAny is null or if not a structure that returns a value, if ScopeDefiningParentIfAny IS a
		/// structure that returns a value (ie. FUNCTION or PROPERTY) then this will be non-null. If this is non-null then ScopeDefiningParentIfAny
		/// will also be non-null.
        /// </summary>
        public CSharpName ParentReturnValueNameIfAny { get; private set; }

        /// <summary>
        /// These are references that are declared as being a compulsory and expected part of the Environment References - eg. if a command line
        /// script is being translated then WScript may be an expected External Dependency and warnings should not be emitted about accessing
        /// it, even though there is nothing to indicate its presence in the source. This will never be null.
        /// </summary>
        public NonNullImmutableList<NameToken> ExternalDependencies { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<ScopedNameToken> Classes { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<ScopedNameToken> Functions { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<ScopedNameToken> Properties { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<ScopedNameToken> Variables { get; private set; }

        public ScopeLocationOptions ScopeLocation
        {
            get
            {
                return (ScopeDefiningParentIfAny == null) ? ScopeLocationOptions.OutermostScope : ScopeDefiningParentIfAny.Scope;
            }
        }
    }
}
