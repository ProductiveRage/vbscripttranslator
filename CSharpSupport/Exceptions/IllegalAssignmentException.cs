using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when a value-setting statement has an invalid target - a constant, for example
    /// </summary>
    [Serializable]
    public class IllegalAssignmentException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Illegal assignment";

        public IllegalAssignmentException(Exception innerException = null) : this(null, innerException) { }
        public IllegalAssignmentException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 501; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected IllegalAssignmentException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
