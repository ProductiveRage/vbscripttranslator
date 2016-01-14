using System;
using System.Runtime.Serialization;

namespace VBScriptTranslator.RuntimeSupport.Exceptions
{
    /// <summary>
    /// This will be raised when a FOR EACH target can not be enumerated over
    /// </summary>
    [Serializable]
    public class ObjectNotCollectionException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object not a collection";

        public ObjectNotCollectionException(Exception innerException = null) : this(null, innerException) { }
        public ObjectNotCollectionException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        public override int ErrorNumber { get { return 451; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected ObjectNotCollectionException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
