using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
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
        public ValueSettingStatement(IEnumerable<IToken> valueToSetTokens, IEnumerable<IToken> expressionTokens, ValueSetTypeOptions valueSetType)
        {
            if (valueToSetTokens == null)
                throw new ArgumentNullException("valueToSetTokens");
            if (expressionTokens == null)
                throw new ArgumentNullException("expressionTokens");
            if (!Enum.IsDefined(typeof(ValueSetTypeOptions), valueSetType))
                throw new ArgumentOutOfRangeException("valueSetType");

            ValueToSetTokens = valueToSetTokens.ToList().AsReadOnly();
            if (!ValueToSetTokens.Any())
                throw new ArgumentException("Empty valueToSetTokens specified - invalid");
            if (ValueToSetTokens.Any(t => t == null))
                throw new ArgumentException("Null token present in valueToSetTokens set");

            ExpressionTokens = expressionTokens.ToList().AsReadOnly();
            if (!ExpressionTokens.Any())
                throw new ArgumentException("Empty expressionTokens specified - invalid");
            if (ExpressionTokens.Any(t => t == null))
                throw new ArgumentException("Null token present in expressionTokens set");

            ValueSetType = valueSetType;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public IEnumerable<IToken> ValueToSetTokens { get; private set; }
        
        public IEnumerable<IToken> ExpressionTokens { get; private set; }
        
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
            var tokensList = ValueToSetTokens.Concat(new[] { AtomToken.GetNewToken("=") }).Concat(ExpressionTokens).ToList();
            if (ValueSetType == ValueSetTypeOptions.Set)
                tokensList.Insert(0, AtomToken.GetNewToken("Set"));

            return (new Statement(tokensList, Statement.CallPrefixOptions.Absent)).GenerateBaseSource(indenter);
        }
    }
}
