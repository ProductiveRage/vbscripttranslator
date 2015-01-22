using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class NumericValueToken : AtomToken
    {
        /// <summary>
        /// The constructor must take the original string content representing the number since it's important to differentiate between "1" and "1.0"
        /// (where the first is declared as an "Integer" in VBScript and the latter as a "Double")
        /// </summary>
        public NumericValueToken(string content, int lineIndex) : base((content ?? "").Trim(), WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            if (string.IsNullOrWhiteSpace(content))
                throw new ArgumentException("Null/blank content specified");

            double numericValue;
            if (!double.TryParse(content, out numericValue))
                throw new ArgumentException("content must be a string representation of a numeric value");

            Value = numericValue;
        }

        /// <summary>
        /// This will never be null or blank, nor have any leading or trailing whitespace. It will always be parseable as a numeric value.
        /// </summary>
        public new string Content { get; private set; }

        public double Value { get; private set; }
    }
}
