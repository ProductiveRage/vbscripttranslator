using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter
{
    public class ScopeAccessInformation
    {
        public ScopeAccessInformation(
            ParentConstructTypeOptions parentConstructType,
            NonNullImmutableList<NameToken> classes,
            NonNullImmutableList<NameToken> functions,
            NonNullImmutableList<NameToken> properties,
            NonNullImmutableList<NameToken> variables)
        {
            if (!Enum.IsDefined(typeof(ParentConstructTypeOptions), parentConstructType))
                throw new ArgumentOutOfRangeException("parentConstructType");
            if (classes == null)
                throw new ArgumentNullException("classes");
            if (functions == null)
                throw new ArgumentNullException("functions");
            if (properties == null)
                throw new ArgumentNullException("properties");
            if (variables == null)
                throw new ArgumentNullException("variables");

            ParentConstructType = parentConstructType;
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
                    ParentConstructTypeOptions.None,
                    new NonNullImmutableList<NameToken>(),
                    new NonNullImmutableList<NameToken>(),
                    new NonNullImmutableList<NameToken>(),
                    new NonNullImmutableList<NameToken>()
                );
            }
        }

        public ParentConstructTypeOptions ParentConstructType { get; private set; }

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
