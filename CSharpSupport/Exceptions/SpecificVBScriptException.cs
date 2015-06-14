using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    [Serializable]
    public abstract class SpecificVBScriptException : Exception
    {
        protected SpecificVBScriptException(string basicErrorDescription, string additionalInformationIfAny, Exception innerException = null)
            : base(GetMessage(basicErrorDescription, additionalInformationIfAny), innerException) { }

        protected SpecificVBScriptException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public abstract int ErrorNumber { get; }
        
        private static string GetMessage(string basicErrorDescription, string additionalInformationIfAny)
        {
            if (basicErrorDescription == null)
                throw new ArgumentNullException("basicErrorDescription");

            var message = basicErrorDescription;
            if (!string.IsNullOrWhiteSpace(additionalInformationIfAny))
                message += ": " + additionalInformationIfAny.Trim();
            return message;
        }
    }
}
