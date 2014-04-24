using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class DimStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public DimStatement(IEnumerable<DimVariable> variables)
        {
            if (variables == null)
                throw new ArgumentNullException("variables");

            Variables = variables.ToList().AsReadOnly();
            if (Variables.Any(v => v == null))
                throw new ArgumentException("Null reference encountered in variables set");
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null nor contain any nulls (though it may be an empty set)
        /// </summary>
        public IEnumerable<DimVariable> Variables { get; private set; }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        public class DimVariable
        {
            public DimVariable(NameToken name, IEnumerable<Expression> dimensions)
            {
                if (name == null)
                    throw new ArgumentNullException("name");
                
                Name = name;
                if (dimensions == null)
                {
                    Dimensions = null;
                    return;
                }
                Dimensions = dimensions.ToList().AsReadOnly();
                if (Dimensions.Any(d => d == null))
                    throw new ArgumentException("Null reference encountered in dimensions set");
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public NameToken Name { get; private set; }

            /// <summary>
            /// Variables list may be null (not explicitly defined as an array), have zero elements (an uninitialised array) or multiple dimensions (but
            /// if the list is non-null and non-empty, it will never contain any null references)
            /// </summary>
            public IEnumerable<Expression> Dimensions { get; private set; }

            public override string ToString()
            {
                return base.ToString() + ":" + Name;
            }
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public virtual string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            var output = new StringBuilder();
            output.Append(indenter.Indent);
            output.Append("Dim ");
            var numberOfVariables = Variables.Count();
            foreach (var indexedVariable in Variables.Select((v, i) => new { Variable = v, Index = i }))
            {
                output.Append(indexedVariable.Variable.Name.Content);
                if (indexedVariable.Variable.Dimensions != null)
                {
                    output.Append("(");
                    var numberOfDimensions = indexedVariable.Variable.Dimensions.Count();
                    foreach (var indexedDimension in indexedVariable.Variable.Dimensions.Select((d, i) => new { Dimension = d, Index = i }))
                    {
                        output.Append(indexedDimension.Dimension.GenerateBaseSource(new NullIndenter()));
                        if (indexedDimension.Index < (numberOfDimensions - 1))
                            output.Append(", ");
                    }
                    output.Append(")");
                }
                if (indexedVariable.Index < (numberOfVariables - 1))
                    output.Append(", ");
            }
            return output.ToString();
        }
    }
}
