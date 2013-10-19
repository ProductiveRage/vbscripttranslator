using System;
using System.Collections.Generic;
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
        private bool isPublic;
        private bool isDefault;
        private NameToken name;
        private List<Parameter> parameters;
        private List<ICodeBlock> statements;
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
            
            this.isPublic = isPublic;
            this.isDefault = isDefault;
            this.name = name;
            this.parameters = parameters;
            this.statements = statements;
        }

        protected abstract string keyWord { get; }

        public override string ToString()
        {
            return base.ToString() + ":" + this.name;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public bool IsPublic
        {
            get { return this.isPublic; }
        }

        public bool IsDefault
        {
            get { return this.isDefault; }
        }

        public NameToken Name
        {
            get { return this.name; }
        }

        public List<Parameter> Parameters
        {
            get { return this.parameters; }
        }

        public List<ICodeBlock> Statements
        {
            get { return this.statements; }
        }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get { return this.statements.AsReadOnly(); }
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
            if (this.isPublic)
                output.Append("Public ");
            else
                output.Append("Private ");
            if (this.isDefault)
                output.Append("Default ");
            output.Append(this.keyWord + " ");
            output.Append(this.name.Content);
            output.Append("(");
            for (int index = 0; index < this.parameters.Count; index++)
            {
                Parameter parameter = this.parameters[index];
                if (parameter.ByRef)
                    output.Append("ByRef ");
                else
                    output.Append("ByVal ");
                output.Append(parameter.Name.Content);
                if (parameter.IsArray)
                    output.Append("()");
                if (index < (this.parameters.Count - 1))
                    output.Append(", ");
            }
            output.AppendLine(")");

            // Render content
            foreach (ICodeBlock block in this.statements)
                output.AppendLine(block.GenerateBaseSource(indenter.Increase()));

            // Close
            output.Append(indenter.Indent + "End " + this.keyWord);
            return output.ToString();
        }
    }
}
