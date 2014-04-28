using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class DimStatement : BaseDimStatement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public DimStatement(IEnumerable<DimVariable> variables) : base(variables)
        {
            if (variables == null)
                throw new ArgumentNullException("variables");

            // Dim statements (like Private and Public class member declarations and unlike ReDim statements) may only have integer constant array
            // dimensions specified, otherwise a compile error will be raised (on that On Error Resume Next can not bury). The integer constant
            // must be zero or greater (-1 is now acceptable, unlike with ReDim).
            if (Variables.Any(v => (v.Dimensions != null) && v.Dimensions.Any(d => !IsValidExpressionForArrayDimension(d))))
                throw new ArgumentException("All array dimensions must be non-negative integer constants unless a ReDim is used");
        }

        private static bool IsValidExpressionForArrayDimension(Expression expression)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");

            var tokens = expression.Tokens.ToArray();
            if (tokens.Length != 1)
                return false;

            var numericValueToken = tokens[0] as NumericValueToken;
            if (numericValueToken == null)
                return false;

            return !numericValueToken.Content.Contains(".") && (numericValueToken.Value >= 0);
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null nor contain any nulls (though it may be an empty set). Any variables that are declared with array dimensions
        /// will only have dimension expressions that consist of single a NumericValueToken, representing non-negative integer values.
        /// </summary>
        public new IEnumerable<DimVariable> Variables { get { return base.Variables; } }
    }
}
