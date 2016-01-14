using System;

namespace VBScriptTranslator.RuntimeSupport
{
    public class DefaultMemberDetails
    {
        private static readonly DefaultMemberDetails _knownToHaveDefaultMemberPresent = new DefaultMemberDetails(true, false, null, null);
        private static readonly DefaultMemberDetails _knownToNotHaveDefaultMemberPresent = new DefaultMemberDetails(false, false, null, null);
        public static DefaultMemberDetails KnownToHaveDefault() { return _knownToHaveDefaultMemberPresent; }
        public static DefaultMemberDetails KnownToNotHaveDefault() { return _knownToNotHaveDefaultMemberPresent; }
        public static DefaultMemberDetails RetrievedResult(object value)
        {
            return new DefaultMemberDetails(true, true, value, null);
        }
        public static DefaultMemberDetails ExceptionWhileEvaluatingDefault(Exception e)
        {
            return new DefaultMemberDetails(true, false, null, e);
        }

        private DefaultMemberDetails(
            bool isDefaultMemberPresent,
            bool wasDefaultMemberRetrieved,
            object defaultMemberValueIfRetrieved,
            Exception exceptionEncounteredWhileEvaluatingDefaultMemberIfAny)
        {
            if ((!isDefaultMemberPresent || wasDefaultMemberRetrieved) && (exceptionEncounteredWhileEvaluatingDefaultMemberIfAny != null))
                throw new ArgumentException("exceptionEncounteredWhileEvaluatingDefaultMemberIfAny may only be non-null if there IS a default member present and it's value failed during retrieval");

            IsDefaultMemberPresent = isDefaultMemberPresent;
            WasDefaultMemberRetrieved = wasDefaultMemberRetrieved;
            DefaultMemberValueIfRetrieved = defaultMemberValueIfRetrieved;
            ExceptionEncounteredWhileEvaluatingDefaultMemberIfAny = exceptionEncounteredWhileEvaluatingDefaultMemberIfAny;
        }

        /// <summary>
        /// TODO
        /// </summary>
        public bool IsDefaultMemberPresent { get; private set; }

        /// <summary>
        /// TODO
        /// </summary>
        public bool WasDefaultMemberRetrieved { get; private set; }

        /// <summary>
        /// This value should be considered undefined if DefaultMemberRetrieved is false (if DefaultMemberRetrieved is true then this may or may not be null,
        /// since null may be a valid value for a default member call)
        /// </summary>
        public object DefaultMemberValueIfRetrieved { get; private set; }

        /// <summary>
        /// This will always be null if IsDefaultMemberPresent is false or if WasDefaultMemberRetrieved is true false and it will also be null if the default value
        /// was successfully evaluated. However, if there IS a default member, but an exception was raised while it was being executed then that exception may need
        /// to be re-thrown after the fact that a default member is present has been processed.
        /// </summary>
        public Exception ExceptionEncounteredWhileEvaluatingDefaultMemberIfAny { get; set; }
    }
}
