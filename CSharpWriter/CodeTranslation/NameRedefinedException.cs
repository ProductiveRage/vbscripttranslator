using System;
using System.Runtime.Serialization;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class NameRedefinedException : Exception
    {
        public NameRedefinedException(NameToken name) : base(GetMessage(name))
        {
            if (name == null)
                throw new ArgumentNullException("name");

            Name = name;
        }

        protected NameRedefinedException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken Name { get; private set; }

        private static string GetMessage(NameToken name)
        {
            if (name == null)
                throw new ArgumentNullException("name");

            return string.Format(
                "Name redefined at line {0}: \"{1}\"",
                name.LineIndex + 1,
                name.Content
            );
        }
    }
}
