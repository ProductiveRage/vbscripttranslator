using System;
using System.Runtime.Serialization;

namespace VBScriptTranslator.RuntimeSupport.Exceptions
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

        public override int ErrorNumber { get { return 9; } } // From http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs241.htm

        protected SubscriptOutOfRangeException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
