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

        public override int ErrorNumber { get { return 13; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected TypeMismatchException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
