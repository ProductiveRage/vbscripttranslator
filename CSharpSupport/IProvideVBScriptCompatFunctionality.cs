using System;
using System.Collections.Generic;

namespace CSharpSupport
{
	public interface IProvideVBScriptCompatFunctionality : IAccessValuesUsingVBScriptRules
	{
        // Arithemetic operators
        double POW(object l, object r);
        double DIV(object l, object r);
        double MULT(object l, object r);
        int INTDIV(object l, object r);
        double MOD(object l, object r);
        double ADD(object l, object r);
        double SUBT(object o);
        double SUBT(object l, object r);

        // String concatenation
        string CONCAT(object l, object r);

        // Logical operators
        int NOT(object o);
        int AND(object l, object r);
        int OR(object l, object r);
        int XOR(object l, object r);

        // Comparison operators
        int EQ(object l, object r);
        int NOTEQ(object l, object r);
        int LT(object l, object r);
        int GT(object l, object r);
        int LTE(object l, object r);
        int GTE(object l, object r);
        int IS(object l, object r);
        int EQV(object l, object r);
        int IMP(object l, object r);

        // Array definitions - TODO: Note that dimensions may be an empty set for NEWARRAY (such that "Dim a()" creates an empty array), though
        // it is not permissible for RESIZEARRAY (since "ReDim b()" is not valid, nor is "ReDim b", there must be at least one dimension). If
        // an empty set of dimensions is specified, the returned reference should be (object[])null, not an empty array such as new object[0].
        // -1 is a valid dimension size, meaning there should be zero elements in that dimension (but the returned returned should be null).
        object NEWARRAY(IEnumerable<object> dimensions);
        object RESIZEARRAY(object array, IEnumerable<object> dimensions);

        // Builtin functions
		/* TODO
			"ERR", // This is NOT a function (it's a value)
			"TIMER", // This IS a function (as are all of the below)
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
		 */

		// TODO: Integration RANDOMIZE functionality
        // TODO: Deal with recording error settings

        void TRAP(Action action, bool enableErrorTrapping = true);

        VBScriptConstants Constants { get; }
    }
}
