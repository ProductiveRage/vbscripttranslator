using System;
using System.Runtime.Serialization;

namespace CSharpSupport.Exceptions
{
    /// <summary>
    /// This will rarely be used because a syntax error in VBScript will commonly be associated with a compile error, which would mean that the script could not
    /// be translated here (an exception would be raised during translation). However, date literals can not always be validated at translation time and will
    /// sometimes need to be checked at runtime - if they contain a month name (eg. #5 Jan 2010#, as opposed to #2010-5-1# or any other solely numeric format)
    /// then their validity may vary depending upon the culture at runtime ("Jan" may be valid when the translation process occurs, if it's run in an English
    /// environment, but if translated program is then executed in an environment where the culture specifies the French language then it won't be valid).
    /// So, in some cases, a syntax error must be raised at runtime - the translated program will be arranged such that, if this is required, the exception
    /// is raised before any other work is processed (which is how it would be with VBScript's interpreter; each request it will re-read the script and
    /// refuse to execute it if any of the literals are invalid for the current environment).
    /// </summary>
    [Serializable]
    public class SyntaxError : SpecificVBScriptException
    {
        private const string BASIC_ERROR_DESCRIPTION = "Syntax error";

        public SyntaxError(Exception innerException = null) : this(null, innerException) { }
        public SyntaxError(string additionalInformationIfAny, Exception innerException = null)
            : base(BASIC_ERROR_DESCRIPTION, additionalInformationIfAny, innerException) { }

        protected SyntaxError(SerializationInfo info, StreamingContext context) : base(info, context) { }
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
    }
}
