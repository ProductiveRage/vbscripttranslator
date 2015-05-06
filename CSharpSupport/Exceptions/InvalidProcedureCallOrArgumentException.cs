using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This when an invalid type of parameter is specified (such as a non-positive startIndex for the INSTR function)
    /// </summary>
    [Serializable]
    public class InvalidProcedureCallOrArgumentException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Invalid procedure call or argument";

        public InvalidProcedureCallOrArgumentException(Exception innerException = null) : this(null, innerException) { }
        public InvalidProcedureCallOrArgumentException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        protected InvalidProcedureCallOrArgumentException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
