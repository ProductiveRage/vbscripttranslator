using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// VBScript has (little-known) support for escaping the names of variables, function and classes so that they can contain almost anything if wrapped in square
    /// brackets. This really is almost ANYTHING, other than closing square brackets and line returns since there is no escape character for this name-wrapping. So
    /// underscores may be used, which are not valid otherwise. So can names that begin with numbers, which are also not otherwise valid. Names may also contain
    /// quotes or whitespace, names may contain ONLY numbers, symbols and/or whitespace. The name itself can be blank since [] is valid. I've never seen this
    /// live in the wild, but it's valid nonetheless and this class is how names of that form are represented.
    /// </summary>
    [Serializable]
    public class EscapedNameToken : IToken
    {
        /// <summary>
        /// The escapedContent value should not include the opening and closing square brackets
        /// </summary>
        /// <param name="escapedContent"></param>
        public EscapedNameToken(string escapedContent)
        {
            // Note that blank or whitespace-only are acceptable for this content so we can only check for null here
            if (escapedContent == null)
                throw new ArgumentNullException("escapedContent");
            if (escapedContent.Contains("]"))
                throw new ArgumentException("escapedContent may not contain a closing square bracket");
            if (escapedContent.Contains("\n"))
                throw new ArgumentException("escapedContent may not contain a line return");

            Content = escapedContent;
        }

        /// <summary>
        /// This will never be null but it may be blank or entirely whitespace (it will never contain line returns or closing square brackets, though)
        /// </summary>
        public string Content { get; private set; }
    }
}
