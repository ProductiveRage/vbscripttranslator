using System;
using System.Runtime.InteropServices;

namespace CSharpSupport
{
    public static class VBScriptConstants
	{
        /// <summary>
        /// For consistency with VBScript, true and false are ints rather than booleans
        /// </summary>
        public static int True { get { return -1; } }

        /// <summary>
        /// For consistency with VBScript, true and false are ints rather than booleans
        /// </summary>
        public static int False { get { return 0; } }

        public static object Empty { get { return null; } }

        /// <summary>
        /// This is what VBScript considers null to be (as opposed to actual null, which is what it consider Empty to be)
        /// </summary>
        public static object Null { get { return DBNull.Value; } }

        /// <summary>
        /// VBScript's Nothing reference is an uninitialised object, not the same as Empty which is an uninitialised value type and not the same as Null
        /// which is DBNull.Value. What this means under the hood is that Nothing is a VARIANT with type VT_EMPTY, Null is a VARIANT with type VT_NULL
        /// and Empty is an uninitialised VARIANT. This means Empty is equivalent to .net's null, VBScript's Null can be mapped to DBNull.Value and
        /// Nothing may be mapped to a DispatchWrapper that wraps null. Note that if this Nothing reference is passed from VBScript into a .net
        /// COM component then it will appear as a .net null. This is expected and consistent behaviour with a Nothing reference generated
        /// natively in VBScript. See http://www.informit.com/articles/article.aspx?p=27219&seqNum=8 for more about the DispatchWrapper.
        /// </summary>
        public static object Nothing { get { return new DispatchWrapper(null); } }

        // VarType Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs57.htm)
        public static int vbEmpty { get { return 0; } }
        public static int vbNull { get { return 1; } }
        public static int vbInteger { get { return 2; } }
        public static int vbLong { get { return 3; } }
        public static int vbSingle { get { return 4; } }
        public static int vbDouble { get { return 5; } }
        public static int vbCurrency { get { return 6; } }
        public static int vbDate { get { return 7; } }
        public static int vbString { get { return 8; } }
        public static int vbObject { get { return 9; } }
        public static int vbError { get { return 10; } }
        public static int vbBoolean { get { return 11; } }
        public static int vbVariant { get { return 12; } }
        public static int vbDataObject { get { return 13; } }
        public static int vbDecimal { get { return 14; } }
        public static int vbByte { get { return 17; } }
        public static int vbArray { get { return 8192; } }

        // MsgBox Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs49.htm) - don't know why these are defined, but they are!
        public static int vbOKOnly { get { return 0; } }
        public static int vbOKCancel { get { return 1; } }
        public static int vbAbortRetryIgnore { get { return 2; } }
        public static int vbYesNoCancel { get { return 3; } }
        public static int vbYesNo { get { return 4; } }
        public static int vbRetryCancel { get { return 5; } }
        public static int vbCritical { get { return 16; } }
        public static int vbQuestion { get { return 32; } }
        public static int vbExclamation { get { return 48; } }
        public static int vbInformation { get { return 64; } }
        public static int vbDefaultButton1 { get { return 0; } }
        public static int vbDefaultButton2 { get { return 256; } }
        public static int vbDefaultButton3 { get { return 512; } }
        public static int vbDefaultButton4 { get { return 768; } }
        public static int vbApplicationModal { get { return 0; } }
        public static int vbSystemModal { get { return 4096; } }
        public static int vbOK { get { return 1; } }
        public static int vbCancel { get { return 2; } }
        public static int vbAbort { get { return 3; } }
        public static int vbRetry { get { return 4; } }
        public static int vbIgnore { get { return 5; } }
        public static int vbYes { get { return 6; } }
        public static int vbNo { get { return 7; } }

        // String Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs53.htm)
        public static char vbCr { get { return '\n'; } }
        public static string vbCrLf { get { return "\r\n"; } }
        public static char vbFormFeed { get { return (char)11; } }
        public static char vbLf { get { return '\n'; } }
        public static string vbNewLine { get { return Environment.NewLine; } }
        public static char vbNullChar { get { return '\0'; } }
        public static string vbNullString { get { return (string)null; } }
        public static char vbTab { get { return '\t'; } }
        public static char vbVerticalTab { get { return (char)11; } }

        // Miscellaneous Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs47.htm)
        public static int vbObjectError { get { return -2147221504; } }

        // Comparison Constants  (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs35.htm)
        public static int vbBinaryCompare { get { return 0; } }
        public static int vbTextCompare { get { return 1; } }

        // Date and Time Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs39.htm)
        public static int vbSunday { get { return 1; } }
        public static int vbMonday { get { return 2; } }
        public static int vbTuesday { get { return 3; } }
        public static int vbWednesday { get { return 4; } }
        public static int vbThursday { get { return 5; } }
        public static int vbFriday { get { return 6; } }
        public static int vbSaturday { get { return 7; } }
        public static int vbFirstJan1 { get { return 1; } }
        public static int vbFirstFourDays { get { return 2; } }
        public static int vbFirstFullWeek { get { return 3; } }
        public static int vbUseSystem { get { return 0; } }
        public static int vbUseSystemDayOfWeek { get { return 0; } }

        // Colour Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs33.htm)
        public static int vbBlack { get { return 0x00; } }
        public static int vbRed { get { return 0xFF; } }
        public static int vbGreen { get { return 0xFF00; } }
        public static int vbYellow { get { return 0xFFFF; } }
        public static int vbBlue { get { return 0xFF0000; } }
        public static int vbMagenta { get { return 0xFF00FF; } }
        public static int vbCyan { get { return 0xFFFF00; } }
        public static int vbWhite { get { return 0xFFFFFF; } }

        // Date Format Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs37.htm)
        public static int vbGeneralDate { get { return 0; } }
        public static int vbLongDate { get { return 1; } }
        public static int vbShortDate { get { return 2; } }
        public static int vbLongTime { get { return 3; } }
        public static int vbShortTime { get { return 4; } }

        /// <summary>
        /// This is the date that VBScript returns for CDate(0)
        /// </summary>
        public static DateTime ZeroDate { get { return new DateTime(1899, 12, 30); } }

        /// <summary>
        /// This is the date that VBScript returns for DateSerial(0, 0, 0)
        /// </summary>
        public static DateTime ZeroDateSerial { get { return new DateTime(1999, 11, 30); } }

        /// <summary>
        /// This is the earliest date that VBScript will represent - equivalent to DateAdd("d", -29, DateAdd("m", -10, DateAdd("yyyy", -1899, DateSerial(0, 0, 0))))
        /// </summary>
        public static DateTime EarliestPossibleDate { get { return new DateTime(100, 1, 1); } }

        /// <summary>
        /// This is the latest date that VBScript will represent - trying to call DateSerial with parameters for a later date or trying to cast a number to
        /// a date that would be later than this date will fail (the first with an "Invalid procedure call or argument" and the latter with an "Overflow")
        /// </summary>
        public static DateTime LatestPossibleDate { get { return new DateTime(9999, 12, 31, 23, 59, 59); } }

        public static Decimal MinCurrencyValue { get { return -922337203685477.5625099999999m; } }
        public static Decimal MaxCurrencyValue { get { return 922337203685477.5625099999999m; } }
    }
}
