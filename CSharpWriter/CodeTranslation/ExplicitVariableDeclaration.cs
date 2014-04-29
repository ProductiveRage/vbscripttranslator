using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class VariableDeclaration
    {
        public VariableDeclaration(NameToken name, VariableDeclarationScopeOptions scope, IEnumerable<uint> constantDimensionsIfAny)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!Enum.IsDefined(typeof(VariableDeclarationScopeOptions), scope))
                throw new ArgumentOutOfRangeException("scope");

            Name = name;
            Scope = scope;
            ConstantDimensionsIfAny = (constantDimensionsIfAny == null) ? null : constantDimensionsIfAny.ToList().AsReadOnly();
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken Name { get; private set; }

        public VariableDeclarationScopeOptions Scope { get; private set; }

        /// <summary>
        /// This will be null if this was not an array declaration and may be an empty set if it is an uninitialised array declaration
        /// (array declarations with specified dimensions will always be non-negative integer constants when Dim, Private or Public is
        /// used, otherwise a VBScript compile error will have been raised - ReDim may be used to specify variable dimensions but they
        /// will be represented by a VariableDeclaration with no dimensions and a separate statement to set the reference to an array)
        /// </summary>
        public IEnumerable<uint> ConstantDimensionsIfAny { get; private set; }
    }
}
