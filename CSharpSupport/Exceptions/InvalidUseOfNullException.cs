using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when VBScript Null is passed in where it isn't accepted (eg. to CDbl)
    /// </summary>
    [Serializable]
    public class InvalidUseOfNullException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Invalid use of null";

        public InvalidUseOfNullException(Exception innerException = null) : this(null, innerException) { }
        public InvalidUseOfNullException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        protected InvalidUseOfNullException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
