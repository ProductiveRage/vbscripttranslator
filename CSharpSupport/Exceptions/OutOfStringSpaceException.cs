using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when a string is required that would be too long (whether that comes from concatenating other strings or by using the STRING method)
    /// </summary>
    [Serializable]
    public class OutOfStringSpaceException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Out of string space";

        public OutOfStringSpaceException(Exception innerException = null) : this(null, innerException) { }
        public OutOfStringSpaceException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 14; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected OutOfStringSpaceException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
