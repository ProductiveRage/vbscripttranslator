 using System;
using System.Runtime.InteropServices;

namespace CSharpSupport
{
    public class VBScriptConstants
	{
		/// <summary>
		/// For consistency with VBScript, true and false are ints rather than booleans
		/// </summary>
		public int True { get { return -1; } }

		/// <summary>
		/// For consistency with VBScript, true and false are ints rather than booleans
		/// </summary>
		public int False { get { return 0; } }

		public object Empty { get { return null; } }
		
		/// <summary>
		/// This is what VBScript considers null to be (as opposed to actual null, which is what it consider Empty to be)
		/// </summary>
		public object Null { get { return DBNull.Value; } }
		
		/// <summary>
		/// VBScript's Nothing reference is an uninitialised object, not the same as Empty which is an uninitialised value type and not the same as Null
		/// which is DBNull.Value. What this means under the hood is that Nothing is a VARIANT with type VT_EMPTY, Null is a VARIANT with type VT_NULL
		/// and Empty is an uninitialised VARIANT. This means Empty is equivalent to .net's null, VBScript's Null can be mapped to DBNull.Value and
		/// Nothing may be mapped to a DispatchWrapper that wraps null. Note that if this Nothing reference is passed from VBScript into a .net
		/// COM component then it will appear as a .net null. This is expected and consistent behaviour with a Nothing reference generated
		/// natively in VBScript.
		/// </summary>
		public object Nothing { get { return new DispatchWrapper(null); } }

		// VarType Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs57.htm)
		public int vbEmpty { get { return 0; } }
		public int vbNull { get { return 1; } }
		public int vbInteger { get { return 2; } }
		public int vbLong { get { return 3; } }
		public int vbSingle { get { return 4; } }
		public int vbDouble { get { return 5; } }
		public int vbCurrency { get { return 6; } }
		public int vbDate { get { return 7; } }
		public int vbString { get { return 8; } }
		public int vbObject { get { return 9; } }
		public int vbError { get { return 10; } }
		public int vbBoolean { get { return 11; } }
		public int vbVariant { get { return 12; } }
		public int vbDataObject { get { return 13; } }
		public int vbDecimal { get { return 14; } }
		public int vbByte { get { return 17; } }
		public int vbArray { get { return 8192; } }

		// MsgBox Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs49.htm) - don't know why these are defined, but they are!
		public int vbOKOnly { get { return 0; } }
		public int vbOKCancel { get { return 1; } }
		public int vbAbortRetryIgnore { get { return 2; } }
		public int vbYesNoCancel { get { return 3; } }
		public int vbYesNo { get { return 4; } }
		public int vbRetryCancel { get { return 5; } }
		public int vbCritical { get { return 16; } }
		public int vbQuestion { get { return 32; } }
		public int vbExclamation { get { return 48; } }
		public int vbInformation { get { return 64; } }
		public int vbDefaultButton1 { get { return 0; } }
		public int vbDefaultButton2 { get { return 256; } }
		public int vbDefaultButton3 { get { return 512; } }
		public int vbDefaultButton4 { get { return 768; } }
		public int vbApplicationModal { get { return 0; } }
		public int vbSystemModal { get { return 4096; } }
		public int vbOK { get { return 1; } }
		public int vbCancel { get { return 2; } }
		public int vbAbort { get { return 3; } }
		public int vbRetry { get { return 4; } }
		public int vbIgnore { get { return 5; } }
		public int vbYes { get { return 6; } }
		public int vbNo { get { return 7; } }

		// String Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs53.htm)
		public char vbCr { get { return '\n'; } }
		public string vbCrLf { get { return "\r\n"; } }
		public char vbFormFeed { get { return (char)11; } }
		public char vbLf { get { return '\n'; } }
		public string vbNewLine { get { return Environment.NewLine; } }
		public char vbNullChar { get { return '\0'; } }
		public string vbNullString { get { return (string)null; } }
		public char vbTab { get { return '\t'; } }
		public char vbVerticalTab { get { return (char)11; } }
		
		// Miscellaneous Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs47.htm)
		public int vbObjectError { get { return -2147221504; } }

		// Comparison Constants  (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs35.htm)
		public int vbBinaryCompare { get { return 0; } }
		public int vbTextCompare { get { return 1; } }
		
		// Date and Time Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs39.htm)
		public int vbSunday { get { return 1; } }
		public int vbMonday { get { return 2; } }
		public int vbTuesday { get { return 3; } }
		public int vbWednesday { get { return 4; } }
		public int vbThursday { get { return 5; } }
		public int vbFriday { get { return 6; } }
		public int vbSaturday { get { return 7; } }
		public int vbFirstJan1 { get { return 1; } }
		public int vbFirstFourDays { get { return 2; } }
		public int vbFirstFullWeek { get { return 3; } }
		public int vbUseSystem { get { return 0; } }
		public int vbUseSystemDayOfWeek { get { return 0; } }

		// Colour Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs33.htm)
		public int vbBlack { get { return 0x00; } }
		public int vbRed { get { return 0xFF; } }
		public int vbGreen { get { return 0xFF00; } }
		public int vbYellow { get { return 0xFFFF; } }
		public int vbBlue { get { return 0xFF0000; } }
		public int vbMagenta { get { return 0xFF00FF; } }
		public int vbCyan { get { return 0xFFFF00; } }
		public int vbWhite { get { return 0xFFFFFF; } }

		// Date Format Constants (http://www.csidata.com/custserv/onlinehelp/vbsdocs/vbs37.htm)
		public int vbGeneralDate { get { return 0; } }
		public int vbLongDate { get { return 1; } }
		public int vbShortDate { get { return 2; } }
		public int vbLongTime { get { return 3; } }
		public int vbShortTime { get { return 4; } }
    }
}
