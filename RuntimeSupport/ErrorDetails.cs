using System;
using VBScriptTranslator.RuntimeSupport.Attributes;

namespace VBScriptTranslator.RuntimeSupport
{
	public class ErrorDetails
	{
		public static ErrorDetails NoError = new ErrorDetails(0, "", "", null);

		public ErrorDetails(int number, string source, string description, Exception originalExceptionIfKnown)
		{
			Number = number;
			Source = source ?? "";
			Description = description ?? "";
			OriginalExceptionIfKnown = originalExceptionIfKnown;
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

		/// <summary>
		/// This will be non-null Number is non-zero and if the exception was caught and translated into an ErrorDetails instance, but it may be null even
		/// if Number is non-zero (if the error was created with a RAISEERROR call)
		/// </summary>
		public Exception OriginalExceptionIfKnown { get; private set; }
	}
}
