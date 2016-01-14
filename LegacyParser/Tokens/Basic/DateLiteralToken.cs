using System;
using VBScriptTranslator.RuntimeSupport;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a single date literal section. It can not be known at this point whether, in all cases, the value is a valid date or not, it may vary depending
    /// upon culture (eg. "1 May 2015" is valid when the English language is used but not for French). Some date literals may be identified as definitely-valid or
    /// definitely-invalid at this point and it is possible to identify those that can not be confirmed until runtime - these will have the RequiresRuntimeValidation
    /// property set to true. Before performing any work, the translated program must verify that all such date literals are indeed valid - this is equivalent to the
    /// VBScript interpreter validating dates when the script is first read and raising a syntax error if any date literal is invalid.
    /// </summary>
    [Serializable]
    public class DateLiteralToken : IToken
    {
        public DateLiteralToken(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            Content = content;
            LineIndex = lineIndex;
            RequiresRuntimeValidation = false;

            // It's not possible to know at translation times whether ANY date literal is valid, since they may rely upon the culture settings when the translated
            // program is run (eg. "1 May 2010" is valid in English but not in French). The translated programs need to support running in different cultures to be
            // consistent with VBScript, where the interpreter reads the script and validates date literals against the current culture on every request. However,
            // it IS possible to fully validate any all-numeric date formats (eg. "1 1", "2015-5-2", "12 99", etc..) and it's possible to identify SOME invalid
            // month-name-containing formats (eg. "May June" or "42 May 99") and where it's NOT possible, we can identify the literal as requiring a runtime
            // check, using whatever culture is in use at that time.
            var limitedDateParser = new DateParser(
                monthNameTranslator: monthName =>
                {
                    RequiresRuntimeValidation = true;
                    return 1;
                },
                defaultYearOverride: 2015 // This value doesn't really matter for this process
            );
            try
            {
                limitedDateParser.Parse(content);
            }
            catch (Exception e)
            {
                throw new ArgumentException("Invalid date format", "content", e);
            }
        }

        /// <summary>
        /// This will not include the hashes in the value, it will never be null or blank. Some date literals may be known at translation time to be valid and some
        /// may require parsing at the runtime of the translated program (if they vary by culture) - the RequiresRuntimeValidation property indicates which this is.
        /// </summary>
        public string Content { get; private set; }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int LineIndex { get; private set; }
        
        public bool RequiresRuntimeValidation { get; private set; }

        public override string ToString()
        {
            return base.ToString() + ":" + Content;
        }
    }
}
