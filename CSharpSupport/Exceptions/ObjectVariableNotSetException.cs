using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This occurs when Nothing is passed in where a VBScript-value-type reference is expected
    /// </summary>
    [Serializable]
    public class ObjectVariableNotSetException : Exception
    {
        public ObjectVariableNotSetException(Exception innerException = null) : base("Object variable not set", innerException) { }
        protected ObjectVariableNotSetException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
