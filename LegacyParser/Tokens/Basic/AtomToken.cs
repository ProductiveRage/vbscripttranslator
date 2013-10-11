using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This token represents a single unprocessed section of script content (not
    /// string or comment) - it is not initialised directly through a constructor,
    /// instead use the static GetNewToken method which try to will return an
    /// appropriate token type (an actual AtomToken, Operator Token or one
    /// of the AbstractEndOfStatement types)
    /// </summary>
    [Serializable]
    public class AtomToken : IToken
    {
        // =======================================================================================
        // CLASS INITIALISATION - INTERNAL
        // =======================================================================================
        protected string content;
        protected AtomToken(string content)
        {
            // Do all this validation AGAIN because we may re-use this from
            // inheriting classes (eg. OperatorToken)
            if (content == null)
                throw new ArgumentNullException("content");
            if (containsWhiteSpace(content))
                throw new ArgumentException("Whitespace encountered in AtomToken - invalid");
            if (content == "")
                throw new ArgumentException("Blank content specified for AtomToken - invalid");
            this.content = content;
        }

        // =======================================================================================
        // CLASS INITIALISATION - PUBLIC
        // =======================================================================================
        /// <summary>
        /// This will return an AtomToken, OperatorToken, EndOfStatementNewLineToken or
        /// EndOfStatementSameLineToken if the content appears valid (it must be non-
        /// null, non-blank and contain no whitespace - unless it's a single line-
        /// return)
        /// </summary>
        public static IToken GetNewToken(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (content == "")
                throw new ArgumentException("Blank content specified for AtomToken - invalid");

            if (content == "\n")
                return new EndOfStatementNewLineToken();
            if (content == ":")
                return new EndOfStatementSameLineToken();

            if (containsWhiteSpace(content))
                throw new ArgumentException("Whitespace in an AtomToken - invalid");

            if (isMustHandleKeyWord(content) || isMiscKeyWord(content))
                return new KeyWordToken(content);
            if (isOperator(content))
                return new OperatorToken(content);
            if (isComparison(content))
                return new ComparisonToken(content);
            if (isMemberAccessor(content))
                return new MemberAccessorOrDecimalPointToken(content);
            if (isArgumentSeparator(content))
                return new ArgumentSeparatorToken(content);
            if (isOpenBrace(content))
                return new OpenBrace(content);
            if (isCloseBrace(content))
                return new CloseBrace(content);

            return new AtomToken(content);
        }

        private static bool containsWhiteSpace(string content)
        {
            if (content == null)
                throw new ArgumentNullException("token");
            string whitespace = " \r\n\t";
            for (int index = 0; index < content.Length; index++)
            {
                if (whitespace.Contains(content.Substring(index, 1)))
                    return true;
            }
            return false;
        }

        // =======================================================================================
        // [PRIVATE] CONTENT DETERMINATION - eg. isOperator
        // =======================================================================================
        /// <summary>
        /// Does the content appear to represent a VBScript operator? An exception will be raised
        /// for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isOperator(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "&", "+", "-", "*", "/", "MOD",
                    "NOT", "AND", "OR", "XOR"
                }
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript comparison? An exception will be raised
        /// for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isComparison(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "<", ">", "=", "IS",
                    "EQV", "IMP"
                }
            );
        }

        protected static bool isMemberAccessor(string atomContent)
        {
            return isType(
                atomContent,
                new string[] { "." }
            );
        }

        protected static bool isArgumentSeparator(string atomContent)
        {
            return isType(
                atomContent,
                new string[] { "," }
            );
        }

        protected static bool isOpenBrace(string atomContent)
        {
            return isType(
                atomContent,
                new string[] { "(" }
            );
        }

        protected static bool isCloseBrace(string atomContent)
        {
            return isType(
                atomContent,
                new string[] { ")" }
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript keyword that will have to be handled by an
        /// AbstractBlockHandler - eg. a "FOR" in a loop, or the "OPTION" from "OPTION EXPLICIT" or the
        /// "RANDOMIZE" command? An exception will be raised for null, blank or whitespace-containing
        /// input.
        /// </summary>
        protected static bool isMustHandleKeyWord(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "OPTION", "EXPLICIT",
                    "DIM", "REDIM", "PRESERVE",
                    "PUBLIC", "PRIVATE",
                    "IF", "THEN", "ELSE", "ELSEIF", "END",
                    "SUB", "FUNCTION", "PROPERTY", "DEFAULT", "CLASS",
                    "SELECT", "CASE", 
                    "FOR", "EACH", "NEXT", "STEP",
                    "DO", "WHILE", "UNTIL", "LOOP", "WEND",
                    "RANDOMIZE",
                    "REM",
                    "GET"
                }
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript keyword that may form part of a general
        /// statement and not have to be handled by a specific AbstractBlockHandler - eg. a "NEW"
        /// declaration for instantiating a class instance. An exception will be raised for null,
        /// blank or whitespace-containing input.
        /// </summary>
        protected static bool isMiscKeyWord(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "CALL",
                    "LET", "SET",
                    "NEW",
                    "ON", "ERROR", "RESUME"
                }
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript expression - eg. "TIMER".
        /// An exception will be raised for null, blank or        /// whitespace-containing input.
        /// </summary>
        private static bool isVBScriptValue(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "TRUE", "FALSE",
                    "EMPTY", "NOTHING", "NULL", 
                    "TIMER",
                    "ERR"
                }
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript function - eg. the "ISNULL" method.
        /// An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        private static bool isVBScriptFunction(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "ISEMPTY", "ISNULL", "ISOBJECT", "ISNUMERIC", "ISDATE", "ISEMPTY", "ISNULL", "ISARRAY",
                    "LBOUND", "UBOUND",
                    "VARTYPE", "TYPENAME",
                    "CREATEOBJECT", "GETOBJECT",
                    "CBYTE", "CINT", "CSNG", "CDBL", "CBOOL", "CSTR", "CDATE",
                    "DATESERIAL", "DATEVALUE", "TIMESERIAL", "TIMEVALUE",
                    "NOW", "DAY", "MONTH", "YEAR", "WEEKDAY", "HOUR", "MINUTE", "SECOND",
                    "ABS", "ATN", "COS", "SIN", "TAN", "EXP", "LOG", "SQR", "RND",
                    "HEX", "OCT", "FIX", "INT", "SNG",
                    "ASC", "ASCB", "ASCW",
                    "CHR", "CHRB", "CHRW",
                    "ASC", "ASCB", "ASCW",
                    "INSTR", "INSTRREV",
                    "LEN", "LENB",
                    "LCASE", "UCASE",
                    "LEFT", "LEFTB", "RIGHT", "RIGHTB", "SPACE",
                    "STRCOMP", "STRING",
                    "LTRIM", "RTRIM", "TRIM"
                }
            );
        }

        private static bool isType(string atomContent, string[] keyWords)
        {
            if (atomContent == null)
                throw new ArgumentNullException("token");
            if (atomContent == "")
                throw new ArgumentException("Blank content specified - invalid");
            if (containsWhiteSpace(atomContent))
                throw new ArgumentException("Whitespace encountered in atomContent - invalid");
            if (keyWords == null)
                throw new ArgumentNullException("keyWords");
            foreach (string keyWord in keyWords)
            {
                if ((keyWord ?? "").Trim() == "")
                    throw new ArgumentException("Null / blank keyWord specified");
                if (containsWhiteSpace(keyWord))
                    throw new ArgumentException("keyWord specified containing whitespce - invalid");
            }
            foreach (string keyWord in keyWords)
            {
                if (atomContent.ToUpper() == keyWord.ToUpper())
                    return true;
            }
            return false;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be blank or null
        /// </summary>
        public string Content
        {
            get { return this.content; }
        }

        /// <summary>
        /// Does this AtomContent describe a reserved VBScript keyword or operator?
        /// </summary>
        public bool IsVBScriptSymbol
        {
            get
            {
                return
                    isMustHandleKeyWord(this.content) ||
                    isMiscKeyWord(this.content) ||
                    isComparison(this.content) ||
                    isOperator(this.content) ||
                    isMemberAccessor(this.content) ||
                    isArgumentSeparator(this.content) ||
                    isOpenBrace(this.content) ||
                    isCloseBrace(this.content) ||
                    isVBScriptFunction(this.content) ||
                    isVBScriptValue(this.content);
            }
        }

        /// <summary>
        /// Does this AtomContent describe a reserved VBScript keyword that must be handled by
        /// a targeted AbstractCodeBlockHandler? (eg. "FOR", "DIM")
        /// </summary>
        public bool IsMustHandleKeyWord
        {
            get { return isMustHandleKeyWord(this.content); }
        }

        /// <summary>
        /// Does this AtomContent describe a VBScript (eg. "ABS")?
        /// </summary>
        public bool IsVBScriptFunction
        {
            get { return isVBScriptFunction(this.content); }
        }

        /// <summary>
        /// Does this AtomContent describe a VBScript value (eg. "TIMER")?
        /// </summary>
        public bool IsVBScriptValue
        {
            get { return isVBScriptValue(this.content); }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + this.content;
        }
    }
}
