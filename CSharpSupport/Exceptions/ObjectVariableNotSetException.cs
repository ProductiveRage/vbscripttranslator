using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when Nothing is passed in where a VBScript-value-type reference is expected
    /// </summary>
    [Serializable]
    public class ObjectVariableNotSetException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object variable not set";

        public ObjectVariableNotSetException(Exception innerException = null) : this(null, innerException) { }
        public ObjectVariableNotSetException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 91; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected ObjectVariableNotSetException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
