using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This token represents a single unprocessed section of script content (not string or comment) - it is not initialised directly through a constructor,
    /// instead use the static GetNewToken method which try to will return an appropriate token type (an actual AtomToken, Operator Token or one of the
    /// AbstractEndOfStatement types)
    /// </summary>
    [Serializable]
    public class AtomToken : IToken
    {
        // =======================================================================================
        // CLASS INITIALISATION - INTERNAL
        // =======================================================================================
        protected string content;
        protected AtomToken(string content, WhiteSpaceBehaviourOptions whiteSpaceBehaviour)
        {
            // Do all this validation AGAIN because we may re-use this from inheriting classes (eg. OperatorToken)
            if (content == null)
                throw new ArgumentNullException("content");
            if (!Enum.IsDefined(typeof(WhiteSpaceBehaviourOptions), whiteSpaceBehaviour))
                throw new ArgumentOutOfRangeException("whiteSpaceBehaviour");
            if ((whiteSpaceBehaviour == WhiteSpaceBehaviourOptions.Disallow) && containsWhiteSpace(content))
                throw new ArgumentException("Whitespace encountered in AtomToken - invalid");
            if (content == "")
                throw new ArgumentException("Blank content specified for AtomToken - invalid");
            this.content = content;
        }

        protected enum WhiteSpaceBehaviourOptions
        {
            Allow,
            Disallow
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

            var recognisedType = TryToGetAsRecognisedType(content);
            if (recognisedType != null)
                return recognisedType;

            if (content.StartsWith("["))
            {
                if (!content.EndsWith("]"))
                    throw new ArgumentException("If content starts with a square bracket then it must have a closing bracket to indicate an escaped-name variable");
                return new EscapedNameToken(content);
            }

            if (containsWhiteSpace(content))
                throw new ArgumentException("Whitespace in an AtomToken - invalid");

            return new NameToken(content);
        }

        /// <summary>
        /// This will try to identify the token content as a VBScript operator or comparison or built-in function or value or line return or statement
        /// separator or numeric value. If unable to match its type then it will return null - this should indicate the name of a function, property,
        /// variable, etc.. defined in the source code being processed.
        /// </summary>
        protected static IToken TryToGetAsRecognisedType(string content)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            if (content == "\n")
                return new EndOfStatementNewLineToken();
            if (content == ":")
                return new EndOfStatementSameLineToken();

            if (isMustHandleKeyWord(content) || isMiscKeyWord(content))
                return new KeyWordToken(content);
            if (isVBScriptFunction(content))
                return new BuiltInFunctionToken(content);
            if (isVBScriptValue(content))
                return new BuiltInValueToken(content);
            if (isLogicalOperator(content))
                return new LogicalOperatorToken(content);
            if (isComparison(content))
                return new ComparisonOperatorToken(content);
            if (isOperator(content))
                return new OperatorToken(content);
            if (isMemberAccessor(content))
                return new MemberAccessorOrDecimalPointToken(content);
            if (isArgumentSeparator(content))
                return new ArgumentSeparatorToken(content);
            if (isOpenBrace(content))
                return new OpenBrace();
            if (isCloseBrace(content))
                return new CloseBrace();

            float numericValue;
            if (float.TryParse(content, out numericValue))
                return new NumericValueToken(numericValue);
			if (content.StartsWith("&h", StringComparison.InvariantCultureIgnoreCase))
			{
				int numericHexValue;
				if (int.TryParse(content.Substring(2), NumberStyles.HexNumber, null, out numericHexValue))
					return new NumericValueToken(numericHexValue);
			}

            return null;
        }

        private static string WhiteSpaceChars = new string(
            Enumerable.Range((int)char.MinValue, (int)char.MaxValue).Select(v => (char)v).Where(c => char.IsWhiteSpace(c)).ToArray()
        );

        protected static bool containsWhiteSpace(string content)
        {
            if (content == null)
                throw new ArgumentNullException("token");

            return content.Any(c => WhiteSpaceChars.IndexOf(c) != -1);
        }

        // =======================================================================================
        // [PRIVATE] CONTENT DETERMINATION - eg. isOperator
        // =======================================================================================
        /// <summary>
        /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
        /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an expression
        /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
        /// </summary>
        public static IEnumerable<string> ArithmeticAndStringOperatorTokenValues = new List<string>
        {
            "^", "/", "*", "\"", "MOD", "+", "-", "&" // Note: "\" is integer division (see the link above)
        }.AsReadOnly();

        /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
        /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an expression
        /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
        public static IEnumerable<string> LogicalOperatorTokenValues = new List<string>
        {
            "NOT", "AND", "OR", "XOR"
        }.AsReadOnly();

        /// <summary>
        /// Does the content appear to represent a VBScript operator (eg. an arithermetic operator such as "*", a logical operator such as "AND" or
        /// a comparison operator such as ">")? An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isOperator(string atomContent)
        {
            return isType(
                atomContent,
                ArithmeticAndStringOperatorTokenValues.Concat(LogicalOperatorTokenValues).Concat(ComparisonTokenValues)
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript operator (eg. AND)? An exception will be raised
        /// for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isLogicalOperator(string atomContent)
        {
            return isType(
                atomContent,
                LogicalOperatorTokenValues
            );
        }

        /// <summary>
        /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
        /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an expression
        /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
        /// </summary>
        public static IEnumerable<string> ComparisonTokenValues = new List<string>
        {
            "=", "<>", "<", ">", "<=", ">=", "IS",
            "EQV", "IMP"
        }.AsReadOnly();

        /// <summary>
        /// Does the content appear to represent a VBScript comparison? An exception will be raised
        /// for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isComparison(string atomContent)
        {
            return isType(
                atomContent,
                ComparisonTokenValues
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
        /// Does the content appear to represent a VBScript expression - eg. "TIMER". An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isVBScriptValue(string atomContent)
        {
            return isType(
                atomContent,
                new string[]
                {
                    "TRUE", "FALSE",
                    "EMPTY", "NOTHING", "NULL", 
                    "TIMER",
                    "ERR",

					// These are the constants from http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbscon3.htm that appear to work in VBScript
					
					// VarType Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs57.htm)
					"vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble", "vbCurrency", "vbDate", "vbString", "vbObject", "vbError", "vbBoolean",
					"vbVariant", "vbDataObject", "vbDecimal", "vbByte", "vbArray",

					// MsgBox Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs49.htm) - don't know why these are defined, but they are!
					"vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel", "vbYesNo", "vbRetryCancel", "vbCritical", "vbQuestion", "vbExclamation",
					"vbInformation", "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3", "vbDefaultButton4", "vbApplicationModal", "vbSystemModal",
					"vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo",
					
					// String Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs53.htm)
					"vbCr", "vbCrLf", "vbFormFeed", "vbLf", "vbNewLine", "vbNullChar", "vbNullString", "vbTab", "vbVerticalTab",
					
					// Miscellaneous Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs47.htm)
					"vbObjectError",
					
					// Comparison Constants  (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs35.htm)
					"vbBinaryCompare", "vbTextCompare",
					
					// Date and Time Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs39.htm)
					"vbSunday", "vbMonday", "vbTuesday", "vbWednesday", "vbThursday", "vbFriday", "vbSaturday", "vbFirstJan1", "vbFirstFourDays", "vbFirstFullWeek",
					"vbUseSystem", "vbUseSystemDayOfWeek",
					
					// Colour Constants ( http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs33.htm)
					"vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue", "vbMagenta", "vbCyan", "vbWhite",
					
					// Date Format Constants ( http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs37.htm)
					"vbGeneralDate", "vbLongDate", "vbShortDate", "vbLongTime", "vbShortTime"
                }
            );
        }

        /// <summary>
        /// Does the content appear to represent a VBScript function - eg. the "ISNULL" method.
        /// An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        protected static bool isVBScriptFunction(string atomContent)
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

        private static bool isType(string atomContent, IEnumerable<string> keyWords)
        {
            if (atomContent == null)
                throw new ArgumentNullException("token");
            if (atomContent == "")
                throw new ArgumentException("Blank content specified - invalid");
            if (containsWhiteSpace(atomContent))
                throw new ArgumentException("Whitespace encountered in atomContent - invalid");
            if (keyWords == null)
                throw new ArgumentNullException("keyWords");
            foreach (var keyWord in keyWords)
            {
                if ((keyWord ?? "").Trim() == "")
                    throw new ArgumentException("Null / blank keyWord specified");
                if (containsWhiteSpace(keyWord))
                    throw new ArgumentException("keyWord specified containing whitespce - invalid");
                if (atomContent.Equals(keyWord, StringComparison.InvariantCultureIgnoreCase))
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
