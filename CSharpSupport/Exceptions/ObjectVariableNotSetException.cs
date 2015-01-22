using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when Nothing is passed in where a VBScript-value-type reference is expected or any object-reference type that does not have a
    /// default function or property that can be retrieved without any arguments
    /// </summary>
    [Serializable]
    public class ObjectVariableNotSetException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Object variable not set";

        public ObjectVariableNotSetException(Exception innerException = null) : this(null, innerException) { }
        public ObjectVariableNotSetException(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        protected ObjectVariableNotSetException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
