using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when Nothing is passed in where a VBScript-value-type reference is expected
    /// </summary>
    [Serializable]
    public class TypeMismatchException : Exception
    {
        public TypeMismatchException(Exception innerException = null) : this(null, innerException) { }
        public TypeMismatchException(string additionalInformationIfAny, Exception innerException = null) : base(GetMessage(additionalInformationIfAny), innerException) { }
        protected TypeMismatchException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }

        private static string GetMessage(string additionalInformationIfAny)
        {
            var message = "Type mismatch";
            if (!string.IsNullOrWhiteSpace(additionalInformationIfAny))
                message += ": " + additionalInformationIfAny.Trim();
            return message;
        }

    }
}
