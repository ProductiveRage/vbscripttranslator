using System;
using System.Collections.Generic;

namespace CSharpSupport
{
    /// <summary>
    /// Note that the SETERROR and CLEARANYERROR methods in this class have no explicit way to be associated with a specific request, so there
    /// should either be instances of implementations of this interface for each request (so that, in a very explicit manner, the error information
    /// from one request can never spread over the error information from another request) or some sort of ThreadLocal or ThreadStatic trickery
    /// may be required to separate this data.
    /// </summary>
	public interface IProvideVBScriptCompatFunctionalityToIndividualRequests : IAccessValuesUsingVBScriptRules
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
        // Some error conditions will result in the target being set to null as well as an error being raised - eg. if a non-numeric string
        // value is passed as one of the dimensions. Some error conditions will not result in the target being changed - eg. if RESIZEARRAY
        // is requested to change the array to give it different dimensions, from (1, 1) to (1, 1, 1), for example. This is to be consistent
        // with VBScript's behaviour.
        void NEWARRAY(IEnumerable<object> dimensions, Action<object> targetSetter);
        void RESIZEARRAY(object array, IEnumerable<object> dimensions, Action<object> targetSetter);

        // Builtin functions - TODO: These are not fully specified yet (eg. LEFT requires more than one parameter and INSTR requires multiple
        // parameters and overloads to deal with optional parameters)
        // - Type conversions
        object CBYTE(object value);
        object CBOOL(object value);
        object CDBL(object value);
        object CDATE(object value);
        object CINT(object value);
        object CLNG(object value);
        object CSNG(object value);
        string CSTR(object value);
        string INT(object value);
        string STRING(object value);
        // - String functions
        object ASC(object value);
        object ASCB(object value);
        object ASCW(object value);
        object CHR(object value);
        object CHRB(object value);
        object CHRW(object value);
        object INSTR(object value);
        object INSTRREV(object value);
        object LEN(object value);
        object LENB(object value);
        object LEFT(object value);
        object LEFTB(object value);
        object RIGHT(object value);
        object RIGHTB(object value);
        object REPLACE(object value);
        object SPACE(object value);
        object SPLIT(object value);
        object STRCOMP(object value);
        string TRIM(object value);
        string LTRIM(object value);
        string RTRIM(object value);
        string LCASE(object value);
        string UCASE(object value);
        // - Type comparisons
        object ISARRAY(object value);
        object ISDATE(object value);
        object ISEMPTY(object value);
        object ISNULL(object value);
        object ISNUMERIC(object value);
        object ISOBJECT(object value);
        object TYPENAME(object value);
        object VARTYPE(object value);
        // - Array functions
        object ARRAY(object value);
        object ERASE(object value);
        object JOIN(object value);
        object LBOUND(object value);
        object UBOUND(object value);
        // - Date functions
        DateTime NOW();
        DateTime DATE();
        DateTime TIME();
        object DATEADD(object value);
        object DATESERIAL(object value);
        object DATEVALUE(object value);
        object TIMESERIAL(object value);
        object TIMEVALUE(object value);
        object NOW(object value);
        object DAY(object value);
        object MONTH(object value);
        object YEAR(object value);
        object WEEKDAY(object value);
        object HOUR(object value);
        object MINUTE(object value);
        object SECOND(object value);
        // - Object creation
        object CREATEOBJECT(object value);
        object GETOBJECT(object value);
        object EVAL(object value);
        object EXECUTE(object value);
        object EXECUTEGLOBAL(object value);

        /* TODO
            "ERR", // This is NOT a function (it's a value)
            "TIMER", // This IS a function (as are all of the below)
         */

		// TODO: Integration RANDOMIZE functionality

        void CLEARANYERROR();
        void SETERROR(int number, string source, string description, bool raiseAsException);

        int GETERRORTRAPPINGTOKEN();
        void RELEASEERRORTRAPPINGTOKEN(int token);

        void STARTERRORTRAPPING(int token);
        void STOPERRORTRAPPING(int token);
        
        void HANDLEERROR(int token, Action action);

        // The error-handling IF behaviour described below sounds unbelievable, but it's true as demonstrated by the script -
        //
        //   On Error Resume Next
        //   If (GetValue()) Then
        //     WScript.Echo "Yes"
        //   End If
        //   Function GetValue()
        //     Err.Raise vbObjectError, "Test", "Test"
        //   End Function
        //
        // This example WILL enter the conditional, even though the GetValue function raises an error and doesn't even try to set its return value
        // to true. An even simplier example is to replace the GetValue() call with 1/0, which will result in a "Division by zero" error if ON
        // ERROR RESUME NEXT is not present, but which will result in both of the above loops being entered if it IS present).
        /// <summary>
        /// This layers error-handling on top of the IAccessValuesUsingVBScriptRules.IF method, if error-handling is enabled for the specified
        /// token then evaluation of the value will be attempted - if an error occurs then it will be recorded and the condition will be treated
        /// as true, since this is VBScript's behaviour. It will throw an exception for a null valueEvaluator or an invalid errorToken.
        /// </summary>
        bool IF(Func<object> valueEvaluator, int errorToken);

        VBScriptConstants Constants { get; }
    }
}
