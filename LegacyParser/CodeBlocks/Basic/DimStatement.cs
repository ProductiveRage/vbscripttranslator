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
            // Dim statements (like Private and Public class member declarations and unlike ReDim statements) may only have integer constant array
            // dimensions specified, otherwise a compile error will be raised (on that On Error Resume Next can not bury). The integer constant
            // must be zero or greater (-1 is now acceptable, unlike with ReDim).
            var constantDimensionArrayVariables = new List<ConstantNonNegativeArrayDimensionDimVariable>();
            foreach (var variable in base.Variables)
            {
                if (variable.Dimensions == null)
                    constantDimensionArrayVariables.Add(new ConstantNonNegativeArrayDimensionDimVariable(variable.Name, null));
                else if (variable.Dimensions.Count() == 0)
                    constantDimensionArrayVariables.Add(new ConstantNonNegativeArrayDimensionDimVariable(variable.Name, new NumericValueToken[0]));
                else
                {
                    var constantDimensions = new List<NumericValueToken>();
                    foreach (var dimension in variable.Dimensions)
                    {
                        var dimensionTokens = dimension.Tokens.ToArray();
                        bool isValidValue;
                        if (dimensionTokens.Length != 1)
                            isValidValue = false;
                        else
                        {
                            var numericValueToken = dimensionTokens[0] as NumericValueToken;
                            if (numericValueToken == null)
                                isValidValue = false;
                            else
                            {
                                if (numericValueToken.Content.Contains(".") || (numericValueToken.Value < 0))
                                    isValidValue = false;
                                else
                                {
                                    constantDimensions.Add(numericValueToken);
                                    isValidValue = true;
                                }
                            }
                        }
                        if (!isValidValue)
                            throw new ArgumentException("All array dimensions must be non-negative integer constants unless a ReDim is used");
                    }
                    constantDimensionArrayVariables.Add(new ConstantNonNegativeArrayDimensionDimVariable(variable.Name, constantDimensions));
                }
            }
            
            // Overwrite the base class' Variables reference with the derived versions of DimVariable
            base.Variables = constantDimensionArrayVariables.ToArray();
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null nor contain any nulls (though it may be an empty set). Any variables that are declared with array dimensions
        /// will only have dimension expressions that consist of single a NumericValueToken, representing non-negative integer values.
        /// </summary>
        public new IEnumerable<ConstantNonNegativeArrayDimensionDimVariable> Variables
        {
            get { return base.Variables.Cast<ConstantNonNegativeArrayDimensionDimVariable>(); }
        }
        
        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        [Serializable]
        public class ConstantNonNegativeArrayDimensionDimVariable : DimVariable
        {
            public ConstantNonNegativeArrayDimensionDimVariable(NameToken name, IEnumerable<NumericValueToken> dimensions)
                : base(name, (dimensions == null) ? null : dimensions.Select(d => new Expression(new[] { d })))
            {
                if (base.Dimensions != null)
                {
                    if (base.Dimensions.Any(d =>
                        (d.Tokens.Count() != 1) ||
                        d.Tokens.Single().Content.Contains('.') ||
                        !(d.Tokens.Single() is NumericValueToken) ||
                        (((NumericValueToken)d.Tokens.Single()).Value < 0)))
                    {
                        throw new ArgumentException("All dimensions must be non-negative integers");
                    }
                }
            }

            /// <summary>
            /// Variables list may be null (not explicitly defined as an array), have zero elements (an uninitialised array) or multiple dimensions (but
            /// if the list is non-null and non-empty, it will never contain any null references nor any values that indicate negative or non-integer
            /// values)
            /// </summary>
            public new IEnumerable<NumericValueToken> Dimensions
            {
                get { return (base.Dimensions == null) ? null : base.Dimensions.Select(d => (NumericValueToken)d.Tokens.Single()); }
            }
        }
    }
}
