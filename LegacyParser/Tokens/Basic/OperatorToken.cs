using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class OperatorToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the same token type while parsing the original content.
        /// </summary>
        public OperatorToken(string content) : base(content)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (content == null)
                throw new ArgumentNullException("content");
            if (content == "")
                throw new ArgumentException("Blank content specified for OperatorToken - invalid");
            if (!AtomToken.isOperator(content))
                throw new ArgumentException("Invalid content specified - not an Operator");
            if (AtomToken.isLogicalOperator(content) && (!(this is LogicalOperatorToken)))
                throw new ArgumentException("This content indicates a LogicalOperatorToken but this instance is not of that type");
            if (AtomToken.isComparison(content) && (!(this is ComparisonOperatorToken)))
                throw new ArgumentException("This content indicates a ComparisonOperatorToken but this instance is not of that type");

            this.content = content;
        }
    }
}
