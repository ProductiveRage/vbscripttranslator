using System;
using System.Linq;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This is used when the source code included Err.Raise - those errors are translated into exceptions of this type
    /// </summary>
    [Serializable]
    public class CustomException : SpecificVBScriptException
    {
        private const string DEFAULT_SOURCE = "(null)";
        private const string DEFAULT_DESCRIPTION = "Unknown runtime error";

        private readonly int _errorNumber;
        public CustomException(int number, string source, string description) : base(GetMessage(source, description), additionalInformationIfAny: null)
        {
            if (number == 0)
                throw new ArgumentOutOfRangeException("number");

            _errorNumber = number;
        }

        /// <summary>
        /// This will never be zero (but it may be a negative or positive value)
        /// </summary>
        public override int ErrorNumber { get { return _errorNumber; } }

        protected CustomException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }

        private static string GetMessage(string source, string description)
        {
            if (string.IsNullOrWhiteSpace(source))
                source = DEFAULT_SOURCE;
            if (string.IsNullOrWhiteSpace(description))
                description = DEFAULT_DESCRIPTION;

            return string.Join(": ", new[] { source, description }.Where(v => !string.IsNullOrWhiteSpace(v)));
        }
    }
}
