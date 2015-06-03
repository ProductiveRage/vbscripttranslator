using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when a non-Object reference is provided where an Object reference is required (eg. with the "IS" comparison)
    /// </summary>
    [Serializable]
    public class ObjectRequiredException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object required";

        public ObjectRequiredException(Exception innerException = null) : this(null, innerException) { }
        public ObjectRequiredException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 424; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected ObjectRequiredException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
