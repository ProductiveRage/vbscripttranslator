using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when a conversion from one type to another is attempted that fails (eg. passing "a" to CDl)
    /// </summary>
    [Serializable]
    public class TypeMismatchException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Type mismatch";

        public TypeMismatchException(Exception innerException = null) : this(null, innerException) { }
        public TypeMismatchException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        protected TypeMismatchException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
