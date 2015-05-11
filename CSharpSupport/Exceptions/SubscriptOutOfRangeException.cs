using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when an invalid array index is requested (eg. if UBOUND(a, 2) is called with a is a one-dimensional array)
    /// </summary>
    [Serializable]
    public class SubscriptOutOfRangeException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Subscript out of range";

        public SubscriptOutOfRangeException(Exception innerException = null) : this(null, innerException) { }
        public SubscriptOutOfRangeException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        protected SubscriptOutOfRangeException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
