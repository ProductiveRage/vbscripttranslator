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
            NonNullImmutableList<NameToken> classes,
            NonNullImmutableList<NameToken> functions,
            NonNullImmutableList<NameToken> properties,
            NonNullImmutableList<NameToken> variables)
        {
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
            Classes = classes;
            Functions = functions;
            Properties = properties;
            Variables = variables;
        }

        public static ScopeAccessInformation Empty
        {
            get
            {
                return new ScopeAccessInformation(
                    null,
					null,
                    null,
                    new NonNullImmutableList<NameToken>(),
                    new NonNullImmutableList<NameToken>(),
                    new NonNullImmutableList<NameToken>(),
                    new NonNullImmutableList<NameToken>()
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
        /// This will be null if there ScopeDefiningParentIfAny is null or if not a structure (ie. FUNCTION or PROPERTY) that returns a value
        /// </summary>
        public CSharpName ParentReturnValueNameIfAny { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> Classes { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> Functions { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> Properties { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> Variables { get; private set; }
    }
}
