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
            NonNullImmutableList<ScopedNameToken> classes,
            NonNullImmutableList<ScopedNameToken> functions,
            NonNullImmutableList<ScopedNameToken> properties,
            NonNullImmutableList<ScopedNameToken> variables,
            ScopeLocationOptions scopeLocation)
        {
            if (classes == null)
                throw new ArgumentNullException("classes");
            if (functions == null)
                throw new ArgumentNullException("functions");
            if (properties == null)
                throw new ArgumentNullException("properties");
            if (variables == null)
                throw new ArgumentNullException("variables");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");

			if ((parentIfAny == null) && (scopeDefiningParentIfAny != null))
				throw new ArgumentException("If scopeDefiningParentIfAny is non-null then parentIfAny must be");

            ParentIfAny = parentIfAny;
			ScopeDefiningParentIfAny = scopeDefiningParentIfAny;
            ParentReturnValueNameIfAny = parentReturnValueNameIfAny;
            Classes = classes;
            Functions = functions;
            Properties = properties;
            Variables = variables;
            ScopeLocation = scopeLocation;
        }

        public static ScopeAccessInformation Empty
        {
            get
            {
                return new ScopeAccessInformation(
                    null,
					null,
                    null,
                    new NonNullImmutableList<ScopedNameToken>(),
                    new NonNullImmutableList<ScopedNameToken>(),
                    new NonNullImmutableList<ScopedNameToken>(),
                    new NonNullImmutableList<ScopedNameToken>(),
                    ScopeLocationOptions.OutermostScope
                );
            }
        }

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

        public ScopeLocationOptions ScopeLocation { get; private set; }
    }
}
