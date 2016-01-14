using VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions;
using VBScriptTranslator.CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation
{
    public class ScopeAccessInformation
    {
        public ScopeAccessInformation(
			IHaveNestedContent parent,
			IDefineScope scopeDefiningParent,
            CSharpName parentReturnValueNameIfAny,
            CSharpName errorRegistrationTokenIfAny,
            DirectedWithReferenceDetails directedWithReferenceIfAny,
            NonNullImmutableList<NameToken> externalDependencies,
            NonNullImmutableList<ScopedNameToken> classes,
            NonNullImmutableList<ScopedNameToken> functions,
            NonNullImmutableList<ScopedNameToken> properties,
            NonNullImmutableList<ScopedNameToken> constants,
            NonNullImmutableList<ScopedNameToken> variables,
            NonNullImmutableList<ExitableNonScopeDefiningConstructDetails> structureExitPoints)
        {
            if (parent == null)
                throw new ArgumentNullException("parent");
            if (scopeDefiningParent == null)
                throw new ArgumentNullException("scopeDefiningParent");
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");
            if (classes == null)
                throw new ArgumentNullException("classes");
            if (functions == null)
                throw new ArgumentNullException("functions");
            if (properties == null)
                throw new ArgumentNullException("properties");
            if (constants == null)
                throw new ArgumentNullException("constants");
            if (variables == null)
                throw new ArgumentNullException("variables");
            if (structureExitPoints == null)
                throw new ArgumentNullException("structureExitPoints");

            Parent = parent;
			ScopeDefiningParent = scopeDefiningParent;
            ParentReturnValueNameIfAny = parentReturnValueNameIfAny;
            ErrorRegistrationTokenIfAny = errorRegistrationTokenIfAny;
            DirectedWithReferenceIfAny = directedWithReferenceIfAny;
            ExternalDependencies = externalDependencies;
            Classes = classes;
            Functions = functions;
            Properties = properties;
            Constants = constants;
            Variables = variables;
            StructureExitPoints = structureExitPoints;
        }

        public static ScopeAccessInformation FromOutermostScope(
            CSharpName outermostScopeWrapperName,
            NonNullImmutableList<ICodeBlock> blocks,
            NonNullImmutableList<NameToken> externalDependencies)
        {
            if (outermostScopeWrapperName == null)
                throw new ArgumentNullException("outermostScopeWrapperName");
            if (blocks == null)
                throw new ArgumentNullException("blocks");
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");

            // This prepares the empty template for the primary instance..
            var outermostScope = new OutermostScope(outermostScopeWrapperName, blocks);
            var initialScope = new ScopeAccessInformation(
                outermostScope, // parent
                outermostScope, // scope-defining parent
                null, // parentReturnValueNameIfAny
                null, // errorRegistrationTokenIfAny
                null, // directedWithReferenceIfAn
                externalDependencies,
                new NonNullImmutableList<ScopedNameToken>(), // classes
                new NonNullImmutableList<ScopedNameToken>(), // functions,
                new NonNullImmutableList<ScopedNameToken>(), // properties,
                new NonNullImmutableList<ScopedNameToken>(), // constants
                new NonNullImmutableList<ScopedNameToken>(), // variables
                new NonNullImmutableList<ExitableNonScopeDefiningConstructDetails>()
            );

            // .. but it needs the classes, functions, properties and variables declared by the code blocks to be registered with it. To do that, this
            // extension method will be used (seems a bit strange since this method is part of the ScopeAccessInformation class and it relies upon an
            // extension method, but that's only because there's no way to define a static extension method, which would have been ideal for this).
            return initialScope.Extend(outermostScope, blocks);
        }

        /// <summary>
        /// This will never be null - if this is a statement within the outermost scope, there should be a construct to identify this (the OuterMostScope
        /// class is intended to be used for this purposes)
        /// </summary>
		public IHaveNestedContent Parent { get; private set; }

		/// <summary>
        /// This will never be null - if this is a statement within the outermost scope, there should be a construct to identify this (the OuterMostScope
        /// class is intended to be used for this purposes). This may be the same reference as Parent or it may be a different one - if, for example this
        /// is for a statement within an IF block then the Parent will be the IF block and and ScopeDefiningParent will be the Function / Property or
        /// OuterMostScope containing the IF.
		/// </summary>
		public IDefineScope ScopeDefiningParent { get; private set; }

        /// <summary>
        /// This will be null if the ScopeDefining is not a structure that returns a value. If ScopeDefiningParent IS a structure that returns a value
        /// (ie. FUNCTION or PROPERTY) then this will be non-null.
        /// </summary>
        public CSharpName ParentReturnValueNameIfAny { get; private set; }

        /// <summary>
        /// This must not be null if error-trapping is to be supported by the current scope. If it is null then error-trapping can be never be applied
        /// to translated statements (it being non-null does not mean that error-trapping will be applied to ALL statements, it depends upon where and
        /// how error-trapping is enabled by the code being translated).
        /// </summary>
        public CSharpName ErrorRegistrationTokenIfAny { get; private set; }

        /// <summary>
        /// This will be non-null if the current scope is within a WITH statement that resolves relative method or property accesses. It will be null
        /// otherwise.
        /// </summary>
        public DirectedWithReferenceDetails DirectedWithReferenceIfAny { get; private set; }

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
        public NonNullImmutableList<ScopedNameToken> Constants { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<ScopedNameToken> Variables { get; private set; }

        /// <summary>
        /// This contains the early-exit variables of containing non-scope-defining structures that may be exited - such as FOR and DO (FUNCTION is not,
        /// for example, one of these types since functions and properties define a new scope). The earlier an entry is, the further up the ancestor
        /// chain it is. This will never be null, but it may be empty (and it may contain multiple entries for the same type - there may be multiple
        /// FOR entries if there are multiple parent FOR structures).
        /// </summary>
        public NonNullImmutableList<ExitableNonScopeDefiningConstructDetails> StructureExitPoints { get; private set; }

        /// <summary>
        /// If the current code is running within a WITH construct then there may be partial references which need to know what to target to complete
        /// the reference.
        /// </summary>
        public class DirectedWithReferenceDetails
        {
            private readonly int _lineIndexOfWithStatement;
            public DirectedWithReferenceDetails(CSharpName referenceName, int lineIndexOfWithStatement)
            {
                if (referenceName == null)
                    throw new ArgumentNullException("referenceName");
                if (lineIndexOfWithStatement < 0)
                    throw new ArgumentOutOfRangeException("lineIndexOfWithStatement");

                ReferenceName = referenceName;
                _lineIndexOfWithStatement = lineIndexOfWithStatement;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public CSharpName ReferenceName { get; private set; }

            public NameToken AsToken()
            {
                return new DoNotRenameNameToken(
                    ReferenceName.Name,
                    _lineIndexOfWithStatement
                );
            }
        }

        public class ExitableNonScopeDefiningConstructDetails
        {
            public ExitableNonScopeDefiningConstructDetails(CSharpName exitEarlyBooleanNameIfAny, ExitableNonScopeDefiningConstructOptions structureType)
            {
                if (!Enum.IsDefined(typeof(ExitableNonScopeDefiningConstructOptions), structureType))
                    throw new ArgumentOutOfRangeException("structureType");

                ExitEarlyBooleanNameIfAny = exitEarlyBooleanNameIfAny;
                StructureType = structureType;
            }
            
            /// <summary>
            /// This will not be null if there are mismatched EXIT DO / FOR statements with DO / FOR loops (an EXIT DO that is within a FOR loop that is
            /// within a DO loop must exit both the translated C# for loop AND the translated C# do loop). If there is no such complicated logic required
            /// then this will be null (though the entry must exist in the StructureExitPoints chain to correctly describe the code arrangement and allow
            /// validation of the source code).
            /// </summary>
            public CSharpName ExitEarlyBooleanNameIfAny { get; private set; }

            public ExitableNonScopeDefiningConstructOptions StructureType { get; private set; }
        }

        public enum ExitableNonScopeDefiningConstructOptions
        {
            Do,
            For
        }
    }
}
