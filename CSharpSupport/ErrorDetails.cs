using VBScriptTranslator.RuntimeSupport.Attributes;

namespace VBScriptTranslator.RuntimeSupport
{
    public class ErrorDetails
    {
        public static ErrorDetails NoError = new ErrorDetails(0, "", "");

        public ErrorDetails(int number, string source, string description)
        {
            Number = number;
            Source = source ?? "";
            Description = description ?? "";
        }

        /// <summary>
        /// This will be zero if there is no current error
        /// </summary>
        [IsDefault]
        public int Number { get; private set; }

        /// <summary>
        /// This will be a blank string if there is no current error
        /// </summary>
        public string Source { get; private set; }

        /// <summary>
        /// This will be a blank string if there is no current error
        /// </summary>
        public string Description { get; private set; }
    }
}
