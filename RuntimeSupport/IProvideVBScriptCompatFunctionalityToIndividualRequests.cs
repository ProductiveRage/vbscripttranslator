﻿using System;
using System.Collections.Generic;

namespace VBScriptTranslator.RuntimeSupport
{
	/// <summary>
	/// Note that the SETERROR and CLEARANYERROR methods in this class have no explicit way to be associated with a specific request, so there
	/// should either be instances of implementations of this interface for each request (so that, in a very explicit manner, the error information
	/// from one request can never spread over the error information from another request) or some sort of ThreadLocal or ThreadStatic trickery
	/// may be required to separate this data.
	/// </summary>
	public interface IProvideVBScriptCompatFunctionalityToIndividualRequests : IAccessValuesUsingVBScriptRules, IDisposable
	{
		// Arithemetic operators
		double POW(object l, object r);
		object DIV(object l, object r);
		object MULT(object l, object r);
		int INTDIV(object l, object r);
		object MOD(object l, object r);
		object ADD(object l, object r);
		object SUBT(object o);
		object SUBT(object l, object r);

		// String concatenation
		object CONCAT(object l, object r);
		/// <summary>
		/// This may never be called with less than two values (otherwise an exception will be thrown)
		/// </summary>
		object CONCAT(params object[] values);

		// Logical operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
		object NOT(object o);
		object AND(object l, object r);
		object OR(object l, object r);
		object XOR(object l, object r);

		// Comparison operators
		/// <summary>
		/// This will return DBNull.Value or boolean value. VBScript has rules about comparisons between "hard-typed" values (aka literals), such
		/// that a comparison between (a = 1) requires that the value "a" be parsed into a numeric value (resulting in a Type Mismatch if this is
		/// not possible). However, this logic must be handled by the translation process before the EQ method is called. Both comparison values
		/// must be treated as non-object-references, so if they are not when passed in then the method will try to retrieve non-object values
		/// from them - if this fails then a Type Mismatch error will be raised (equivalent to calling VAL for both values). If there are no
		/// issues in preparing both comparison values, this will return DBNull.Value if either value is DBNull.Value and a boolean otherwise.
		/// </summary>
		object EQ(object l, object r);
		object NOTEQ(object l, object r);
		object LT(object l, object r);
		object GT(object l, object r);
		object LTE(object l, object r);
		object GTE(object l, object r);
		bool IS(object l, object r);
		object EQV(object l, object r);
		object IMP(object l, object r);
		/// <summary>
		/// This takes the logic from LT but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which LT has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		bool StrictLT(object l, object r);
		/// <summary>
		/// This takes the logic from LTE but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which LTE has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		bool StrictLTE(object l, object r);
		/// <summary>
		/// This takes the logic from GT but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which GT has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		bool StrictGT(object l, object r);
		/// <summary>
		/// This takes the logic from GTE but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which GTE has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		bool StrictGTE(object l, object r);

		/// <summary>
		/// This returns the value without any immediate processing, but may keep a reference to it and dispose of it (where applicable) after
		/// the request completes (to try to avoid resources from not being cleaned up in the absence of the VBScript deterministic garbage
		/// collection - classes with a Class_Terminate function are translated into IDisposable types and, while IDisposable.Dispose will not
		/// be called by the translated code, it may be called after the request ends if the requests are tracked here. This will throw an
		/// exception for a null value.
		/// </summary>
		object NEW(object value);

		// Note that dimensions may be an empty set for NEWARRAY (such that "Dim a()" creates an empty array), though it is not permissible for
		// RESIZEARRAY (since "ReDim b()" is not valid, nor is "ReDim b", there must be at least one dimension). If an empty set of dimensions
		// is specified, the returned reference should be (object[])null, not an empty array such as new object[0]. -1 is a valid dimension size,
		// meaning there should be zero elements in that dimension (but the returned returned should be null). Some error conditions will result
		// in the target being set to null as well as an error being raised - eg. if a non-numeric string value is passed as one of the dimensions.
		// Some error conditions will not result in the target being changed - eg. if RESIZEARRAY is requested to change the array to give it
		// different dimensions, from (1, 1) to (1, 1, 1), for example. This is to be consistent with VBScript's behaviour.
		object NEWARRAY(IEnumerable<object> dimensions);
		object RESIZEARRAY(object array, IEnumerable<object> dimensions);

		// Builtin functions - TODO: These are not fully specified yet (eg. LEFT requires more than one parameter and INSTR requires multiple
		// parameters and overloads to deal with optional parameters)
		// - Type conversions
		byte CBYTE(object value);
		bool CBOOL(object value);
		decimal CCUR(object value);
		double CDBL(object value);
		DateTime CDATE(object value);
		Int16 CINT(object value);
		int CLNG(object value);
		float CSNG(object value);
		string CSTR(object value);
		object INT(object value);
		string STRING(object numberOfTimesToRepeat, object character);

		// - Number functions
		object ABS(object value);
		object ATN(object value);
		object COS(object value);
		object EXP(object value);
		object FIX(object value);
		object LOG(object value);
		object OCT(object value);
		float RND();
		float RND(object value);
		object ROUND(object value); // TODO: See http://blogs.msdn.com/b/ericlippert/archive/2003/09/26/bankers-rounding.aspx
		object SGN(object value);
		object SIN(object value);
		object SQR(object value);
		object TAN(object value);
		// - String functions
		short ASC(object value);
		object ASCB(object value);
		short ASCW(object value);
		string CHR(object value);
		object CHRB(object value);
		object CHRW(object value);
		object ESCAPE(object value);
		object FORMATCURRENCY(object value);
		object FORMATDATETIME(object value);
		object FORMATNUMBER(object value); // TODO: See http://blogs.msdn.com/b/ericlippert/archive/2003/09/26/53112.aspx
		object FORMATPERCENT(object value);
		object FILTER(object value);
		object HEX(object value);
		object INSTR(object valueToSearch, object valueToSearchFor);
		object INSTR(object startIndex, object valueToSearch, object valueToSearchFor);
		object INSTR(object startIndex, object valueToSearch, object valueToSearchFor, object compareMode);
		object INSTRREV(object valueToSearch, object valueToSearchFor);
		object INSTRREV(object valueToSearch, object valueToSearchFor, object startIndex);
		object INSTRREV(object valueToSearch, object valueToSearchFor, object startIndex, object compareMode);
		object LEN(object value);
		object LENB(object value);
		object LEFT(object value, object maxLength);
		object LEFTB(object value, object maxLength);
		object MID(object value, object start);
		object MID(object value, object start, object length);
		object RGB(object value);
		object RIGHT(object value, object maxLength);
		object RIGHTB(object value, object maxLength);
		string REPLACE(object value, object toSearchFor, object toReplaceWith);
		string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex);
		string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements);
		string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode);
		object SPACE(object value);
		object[] SPLIT(object value);
		object[] SPLIT(object value, object delimiter);
		object STRCOMP(object string1, object string2);
		object STRCOMP(object string1, object string2, object compare);
		object STRREVERSE(object value);
		object TRIM(object value);
		object LTRIM(object value);
		object RTRIM(object value);
		object LCASE(object value);
		object UNESCAPE(object value);
		object UCASE(object value);
		// - Type comparisons
		bool ISARRAY(object value);
		bool ISDATE(object value);
		bool ISEMPTY(object value);
		bool ISNULL(object value);
		bool ISNUMERIC(object value);
		bool ISOBJECT(object value);
		string TYPENAME(object value);
		object VARTYPE(object value);
		// - Array functions
		object ARRAY(params object[] value);
		void ERASE(object target, Action<object> targetSetter);
		void ERASE(object target, params object[] arguments);
		string JOIN(object value);
		string JOIN(object value, object delimeter);
		int LBOUND(object value, object dimension);
		int LBOUND(object value);
		int UBOUND(object value);
		int UBOUND(object value, object dimension);
		// - Date functions
		DateTime NOW();
		DateTime DATE();
		DateTime TIME();
		object DATEADD(object interval, object number, object value);
		object DATEDIFF(object interval, object date1, object date2);
		object DATEPART(object value);
		object DATESERIAL(object year, object month, object date);
		DateTime DATEVALUE(object value);
		object DAY(object value);
		object MONTH(object value);
		object MONTHNAME(object value);
		object YEAR(object value);
		object WEEKDAY(object value);
		object WEEKDAY(object value, object firstDayOfWeek);
		object WEEKDAYNAME(object value);
		object WEEKDAYNAME(object value, object abbreviate);
		object WEEKDAYNAME(object value, object abbreviate, object firstDayOfWeek);
		object HOUR(object value);
		object MINUTE(object value);
		object SECOND(object value);
		DateTime TIMESERIAL(object hours, object minutes, object seconds);
		DateTime TIMEVALUE(object value);
		// - Object creation
		object CREATEOBJECT(object value);
		object GETOBJECT(object value);
		object EVAL(object value);
		object EXECUTE(object value);
		object EXECUTEGLOBAL(object value);
		// - Misc
		object GETLOCALE(object value);
		object GETREF(object value);
		object INPUTBOX(object value);
		object LOADPICTURE(object value);
		object MSGBOX(object value);
		string SCRIPTENGINE(object value);
		int SCRIPTENGINEBUILDVERSION(object value);
		int SCRIPTENGINEMAJORVERSION(object value);
		int SCRIPTENGINEMINORVERSION(object value);
		object SETLOCALE(object value);

		object NEWREGEXP(); // TODO

		/// <summary>
		/// This will never be null (if there is no error then an ErrorDetails with Number zero will be returned)
		/// </summary>
		ErrorDetails ERR { get; }

		/// <summary>
		/// Returns the number of seconds that have elapsed since midnight
		/// </summary>
		Single TIMER();

		void RANDOMIZE();
		void RANDOMIZE(object seed);

		/// <summary>
		/// There are some occassions when the translated code needs to throw a runtime exception based on the content of the source code - eg.
		///   WScript.Echo 1()
		/// It is clear that "1" is a numeric constant and not a function, and so may not be called as one. However, this is not invalid VBScript and so is
		/// not a compile time error, it is something that must result in an exception at runtime. In these cases, where it is known at the time of translation
		/// that an exception must be thrown, this method may be used to do so at runtime. This is different to SETERROR, since that records an exception that
		/// has already been thrown - this throws the specified exception (it returns an object, rather than void, for the same reason as the below signatures).
		/// </summary>
		object RAISEERROR(Exception e);

		// These method signatures have to return a value since these are what are called when the source code includes "Err.Raise 123", which VBScript allows
		// to exist in the form "If (Err.Raise(123)) Then" - if these didn't return values then there could be compile errors in the generated C# that were
		// valid VBScript.
		object RAISEERROR(object number);
		object RAISEERROR(object number, object source);
		object RAISEERROR(object number, object source, object description);

		void CLEARANYERROR();
		void SETERROR(Exception e);

		int GETERRORTRAPPINGTOKEN();
		void RELEASEERRORTRAPPINGTOKEN(int token);

		void STARTERRORTRAPPINGANDCLEARANYERROR(int token);
		void STOPERRORTRAPPINGANDCLEARANYERROR(int token);
		
		// If this allows an error to be raised (ie. the error token does not have error-trapping currently enabled) then the token is then implicitly
		// released (so RELEASEERRORTRAPPINGTOKEN must be called at scope termination points if no errors occur or if they are all trapped, but if
		// there is an early exit due to an error not being caught, then the token will still be released)
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

		/// <summary>
		/// Where date literals were present in the source code, in a format that does not specify a date, they must be translated into dates at
		/// runtime. They must all be expanded to have whatever year it was when the request started - if the request happens to take sufficient
		/// time that the year ticks over during processing, all date literals (without explicit years) must be associated with the year when
		/// the request started. Note that if a new request starts with a different year, then date literals without years within that request
		/// must be associated with the new year (this is consistent with how the VBScript interpreter would re-process the script each time).
		/// </summary>
		DateParser DateLiteralParser { get; }
	}
}
