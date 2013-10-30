using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ValueSettingStatement : ICodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// This statement represents the setting of one value to the result of another expression, whether that be a fixed
        /// value, a variable's value or the return value of a method call
        /// </summary>
        public ValueSettingStatement(Expression valueToSet, Expression expression, ValueSetTypeOptions valueSetType)
        {
            if (valueToSet == null)
                throw new ArgumentNullException("valueToSet");
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (!Enum.IsDefined(typeof(ValueSetTypeOptions), valueSetType))
                throw new ArgumentOutOfRangeException("valueSetType");

			ValueToSet = valueToSet;
			Expression = expression;
			ValueSetType = valueSetType;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
		/// <summary>
		/// This will never be null
		/// </summary>
		public Expression ValueToSet { get; private set; }

		/// <summary>
		/// This will never be null
		/// </summary>
		public Expression Expression { get; private set; }
        
        public ValueSetTypeOptions ValueSetType { get; private set; }

        public enum ValueSetTypeOptions
        {
            Let,
            Set
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
            // The Statement class' GenerateBaseSource has logic about rendering strings of tokens and rules about whitespace around
            // (or not around) particular tokens, so the content from this class is wrapped up as a Statement so that the method may
            // be re-used without copying any of it here
            var assignmentOperator = AtomToken.GetNewToken("=", ValueToSet.Tokens.Last().LineIndex);
            var tokensList = ValueToSet.Tokens.Concat(new[] { assignmentOperator }).Concat(Expression.Tokens).ToList();
            if (ValueSetType == ValueSetTypeOptions.Set)
                tokensList.Insert(0, AtomToken.GetNewToken("Set", ValueToSet.Tokens.First().LineIndex));

            return (new Statement(tokensList, Statement.CallPrefixOptions.Absent)).GenerateBaseSource(indenter);
        }
    }
}
