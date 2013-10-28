using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public abstract class AbstractFunctionBlock : ICodeBlock, IDefineScope
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public AbstractFunctionBlock(
            bool isPublic,
            bool isDefault,
            NameToken name,
            List<Parameter> parameters,
            List<ICodeBlock> statements)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!isPublic && isDefault)
                throw new ArgumentException("Can not be default AND private");
            if (parameters == null)
                throw new ArgumentNullException("parameters");
            if (statements == null)
                throw new ArgumentNullException("statements");

            Parameters = parameters.ToList().AsReadOnly();
            if (Parameters.Any(p => p == null))
                throw new ArgumentException("Null reference encountered in parameters set");
            Statements = statements.ToList().AsReadOnly();
            if (Statements.Any(s => s == null))
                throw new ArgumentException("Null reference encountered in Statements set");
            
            IsPublic = isPublic;
            IsDefault = isDefault;
            Name = name;
        }

        protected abstract string keyWord { get; }

        public override string ToString()
        {
            return base.ToString() + ":" + Name.Content;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public bool IsPublic { get; private set; }

        public bool IsDefault { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken Name { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references
        /// </summary>
        public IEnumerable<Parameter> Parameters { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references
        /// </summary>
        public IEnumerable<ICodeBlock> Statements { get; private set; }

        /// <summary>
        /// This must never be null but it may be empty (this may be the names of a a function's arguments, for example)
        /// </summary>
        IEnumerable<NameToken> IDefineScope.ExplicitScopeAdditions { get { return Parameters.Select(p => p.Name); } }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get { return Statements; }
        }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        public class Parameter
        {
            private bool byRef;
            private NameToken name;
            private bool isArray;
            public Parameter(bool byRef, NameToken name, bool isArray)
            {
                if (name == null)
                    throw new ArgumentNullException("name");

                this.byRef = byRef;
                this.name = name;
                this.isArray = isArray;
            }

            public bool ByRef { get { return this.byRef; } }
            
            /// <summary>
            /// This will never be null
            /// </summary>
            public NameToken Name { get { return this.name; } }
            
            public bool IsArray { get { return this.isArray; } }
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            // Ensure derived class has behaved itself
            if ((this.keyWord ?? "").Trim() == "")
                throw new Exception("Derived class has not defined non-blank/null keyWord");

            // Render opening declaration (scope, name, arguments)
            StringBuilder output = new StringBuilder();
            output.Append(indenter.Indent);
            if (IsPublic)
                output.Append("Public ");
            else
                output.Append("Private ");
            if (IsDefault)
                output.Append("Default ");
            output.Append(this.keyWord + " ");
            output.Append(Name.Content);
            output.Append("(");
            var numberOfParameters = Parameters.Count();
            for (int index = 0; index < numberOfParameters; index++)
            {
                Parameter parameter = Parameters.ElementAt(index);
                if (parameter.ByRef)
                    output.Append("ByRef ");
                else
                    output.Append("ByVal ");
                output.Append(parameter.Name.Content);
                if (parameter.IsArray)
                    output.Append("()");
                if (index < (numberOfParameters - 1))
                    output.Append(", ");
            }
            output.AppendLine(")");

            // Render content
            foreach (var block in Statements)
                output.AppendLine(block.GenerateBaseSource(indenter.Increase()));

            // Close
            output.Append(indenter.Indent + "End " + this.keyWord);
            return output.ToString();
        }
    }
}
