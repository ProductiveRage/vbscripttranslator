using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This is used whenever VBScript would overflow at runtime (using this instead of the .net OverflowException will more easily enable error code mappings
    /// if that functionality is added - where we ensure that any meta data about VBScript exceptions is consistent with the interpreter, such as Err.Number)
    /// </summary>
    [Serializable]
    public class VBScriptOverflowException : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Overflow";

        public VBScriptOverflowException(double value, Exception innerException = null) : this(value.ToString(), innerException) { }
        public VBScriptOverflowException(decimal value, Exception innerException = null) : this(value.ToString(), innerException) { }
        
        private VBScriptOverflowException(string valueAsString, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, "'[number: " + valueAsString + "]'", innerException) { }

        protected VBScriptOverflowException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
