using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using VBScriptTranslator.RuntimeSupport.Attributes;
using VBScriptTranslator.RuntimeSupport.Exceptions;

namespace VBScriptTranslator.RuntimeSupport.Implementations
{
	/// <summary>
	/// Instances of this class should be used only by a single request and so is not written to be thread safe. This is partly because the SETERROR and
	/// CLEARANYERROR methods have no explicit way to be associated with a specific request (which is not a problem if each instance is associated with
	/// one specific request) but also so that it can be explicitly disposed after each request completes to ensure that any unmanaged resources are
	/// cleaned up. VBScript's deterministic garbage collector can tidy up more aggressively, relying upon reference counting, the best that we can do
	/// with the C# code is for this to implement IDiposable and to ensure that everything is tidy when the request completes and Dispose is called.
	/// </summary>
	public class DefaultRuntimeFunctionalityProvider : IProvideVBScriptCompatFunctionalityToIndividualRequests
	{
		/// <summary>
		/// VBScript has a string length limited by its data storage mechanism; each character is represented by two bytes and the index into that
		/// array of data must be an signed int, since it is capped at half of int.MaxValue.. minus one. I'm not sure if the minus one is to do with
		/// a requirement for there to be a null terminator at the end or some VBScript one-based-index weirdness.. or something else.
		/// </summary>
		private static readonly int MAX_VBSCRIPT_STRING_LENGTH = (int.MaxValue / 2) - 1;

		private readonly IAccessValuesUsingVBScriptRules _valueRetriever;
		private readonly List<IDisposable> _disposableReferencesToClearAfterTheRequest;
		private readonly Queue<int> _availableErrorTokens;
		private readonly Dictionary<int, ErrorTokenState> _activeErrorTokens;
		private readonly DefaultArithmeticFunctionalityProvider _arithmeticHandler;
		private Exception _trappedErrorIfAny;
		public DefaultRuntimeFunctionalityProvider(Func<string, string> nameRewriter, IAccessValuesUsingVBScriptRules valueRetriever)
		{
			if (valueRetriever == null)
				throw new ArgumentNullException("valueRetriever");

			_valueRetriever = valueRetriever;
			_disposableReferencesToClearAfterTheRequest = new List<IDisposable>();
			_availableErrorTokens = new Queue<int>();
			_activeErrorTokens = new Dictionary<int, ErrorTokenState>();
			_arithmeticHandler = new DefaultArithmeticFunctionalityProvider(valueRetriever);
			DateLiteralParser = new DateParser(DateParser.DefaultMonthNameTranslator, DateTime.Now.Year);
			_trappedErrorIfAny = null;
		}

		private enum ErrorTokenState
		{
			OnErrorResumeNext,
			OnErrorGoto0
		}

		~DefaultRuntimeFunctionalityProvider()
		{
			Dispose(false);
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			if (disposing)
			{
				foreach (var disposableResource in _disposableReferencesToClearAfterTheRequest)
				{
					try { disposableResource.Dispose(); }
					catch { }
				}
			}
		}

		// Arithemetic operators
		public object ADD(object l, object r) { return _arithmeticHandler.ADD(l, r); }
		public object SUBT(object o) { return _arithmeticHandler.SUBT(o); }
		public object SUBT(object l, object r) { return _arithmeticHandler.SUBT(l, r); }
		public object MULT(object l, object r) { return _arithmeticHandler.MULT(l, r); }
		public object DIV(object l, object r) { return _arithmeticHandler.DIV(l, r); }
		public int INTDIV(object l, object r) { return _arithmeticHandler.INTDIV(l, r); }
		public double POW(object l, object r) { return _arithmeticHandler.POW(l, r); }
		public object MOD(object l, object r) { return _arithmeticHandler.MOD(l, r); }

		// String concatenation
		public object CONCAT(object l, object r)
		{
			// Try to get both values as value types - if either is Nothing or an object without a default parameterless member, then it's an ObjectVariableNotSetException
			// or ObjectDoesNotSupportPropertyOrMemberException, resp. If one is an object WITH a default parameterless member, but the value of that member is not a value
			// type, then it's a TypeMismatchException. (If the values are value types to begin with, or are objects with default parameterless member that has a value
			// type, then there's nothing to worry about).
			bool parameterLessDefaultMemberWasAvailable;
			if (!TryVAL(l, out parameterLessDefaultMemberWasAvailable, out l))
			{
				if (parameterLessDefaultMemberWasAvailable)
					throw new TypeMismatchException();
				if (IsVBScriptNothing(l))
					throw new ObjectVariableNotSetException();
				throw new ObjectDoesNotSupportPropertyOrMemberException();
			}
			if (!TryVAL(r, out parameterLessDefaultMemberWasAvailable, out r))
			{
				if (parameterLessDefaultMemberWasAvailable)
					throw new TypeMismatchException();
				if (IsVBScriptNothing(r))
					throw new ObjectVariableNotSetException();
				throw new ObjectDoesNotSupportPropertyOrMemberException();
			}
			if ((l == DBNull.Value) && (r == DBNull.Value))
				return DBNull.Value;
			var lString = (l == DBNull.Value) ? "" : _valueRetriever.STR(l);
			var rString = (r == DBNull.Value) ? "" : _valueRetriever.STR(r);
			if ((lString.Length + rString.Length) > MAX_VBSCRIPT_STRING_LENGTH)
				throw new OutOfStringSpaceException();
			return lString + rString;
		}

		/// <summary>
		/// This may never be called with less than two values (otherwise an exception will be thrown)
		/// </summary>
		public object CONCAT(params object[] values)
		{
			if (values == null)
				throw new ArgumentNullException("values");

			if (values.Length < 2)
				throw new ArgumentException("There must be at least two values specified for the CONCAT operation");

			// Concatenate the first two values (using the standard two-value version of the method) and then concatenate each further values on to
			// this accumulator. This could very likely be done in a more efficient manner by recursively splitting the array of values but this will
			// do for now.
			var combinedValue = CONCAT(values[0], values[1]);
			foreach (var additionalValue in values.Skip(2))
				combinedValue = CONCAT(combinedValue, additionalValue);
			return combinedValue;
		}

		// Logical operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
		// - Read http://blogs.msdn.com/b/ericlippert/archive/2004/07/15/184431.aspx
		public object NOT(object o)
		{
			var bitwiseOperationValues = GetForBitwiseOperations("'Not'", o);
			var valueToNot = bitwiseOperationValues.Item1.Single();
			if (valueToNot == null)
			{
				// GetForBitwiseOperations returns nullable int values - since VBScript's Empty (ie. C#'s null) will be interpreted as zero then any
				// null values here mean VBScript's null (ie. DBNull.Value), and so that is what must be returned from this function
				return DBNull.Value;
			}
			return bitwiseOperationValues.Item2(~valueToNot.Value); // Note: VBScript's Not operation is bitwise, not logical (so the ~ operator is used)
		}
		public object AND(object l, object r)
		{
			var bitwiseOperationValues = GetForBitwiseOperations("'And'", l, r);
			var left = bitwiseOperationValues.Item1.First();
			var right = bitwiseOperationValues.Item1.Skip(1).Single();
			if ((left == null) || (right == null))
			{
				// If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When AND'ing, if either (or both)
				// values are Null then Null is returned.
				return DBNull.Value;
			}
			return bitwiseOperationValues.Item2(left.Value & right.Value);
		}
		public object OR(object l, object r)
		{
			var bitwiseOperationValues = GetForBitwiseOperations("'Or'", l, r);
			var left = bitwiseOperationValues.Item1.First();
			var right = bitwiseOperationValues.Item1.Skip(1).Single();
			if ((left == null) && (right == null))
			{
				// If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When OR'ing, if one value is Null
				// but the other isn't, then the non-Null value is returned - only if both values are Null is Null returned.
				return DBNull.Value;
			}
			else if (left == null)
				return right;
			else if (right == null)
				return left;
			return bitwiseOperationValues.Item2(left.Value | right.Value);
		}
		public object XOR(object l, object r) { throw new NotImplementedException(); }

		// Comparison operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
		/// <summary>
		/// This will return DBNull.Value or boolean value. VBScript has rules about comparisons between "hard-typed" values (aka literals), such
		/// that a comparison between (a = 1) requires that the value "a" be parsed into a numeric value (resulting in a Type Mismatch if this is
		/// not possible). However, this logic must be handled by the translation process before the EQ method is called. Both comparison values
		/// must be treated as non-object-references, so if they are not when passed in then the method will try to retrieve non-object values
		/// from them - if this fails then a Type Mismatch error will be raised. If there are no issues in preparing both comparison values,
		/// this will return DBNull.Value if either value is DBNull.Value and a boolean otherwise.
		/// </summary>
		public object EQ(object l, object r) { return ToVBScriptNullable(EQ_Internal(l, r)); }
		private bool? EQ_Internal(object l, object r)
		{
			// Both sides of the comparison must be simple VBScript values (ie. not object references) - pushing both values through VAL will handle
			// that (an exception will be raised if this operation fails and the value will not be affect if it was already an acceptable type)
			l = _valueRetriever.VAL(l);
			r = _valueRetriever.VAL(r);

			// Let's get the outliers out of the way; VBScript Null and Empty..
			if ((l == DBNull.Value) || (r == DBNull.Value))
				return null; // If one or both sides of the comparison are "Null" then this is what is returned
			if ((l == null) && (r == null))
				return true; // If both sides are Empty then they are considered to match
			else if ((l == null) || (r == null))
			{
				// The default values of VBScript primitives (number, strings and booleans) are considered to match Empty
				var nonNullValue = l ?? r;
				if ((IsDotNetNumericType(nonNullValue) && (Convert.ToDouble(nonNullValue)) == 0)
				|| ((nonNullValue as string) == "")
				|| ((nonNullValue is bool) && !(bool)nonNullValue))
					return true;
				return false;
			}

			// Booleans have some funny behaviour in that they will match values of other types (numbers, but not strings unless string literals
			// are in the comparison, which is not logic that this method has to deal with). If one of the values is a boolean and the other isn't,
			// and none of the special cases are met, then there must not be a match.
			if ((l is bool) && (r is bool))
				return (bool)l == (bool)r;
			else if ((l is bool) || (r is bool))
			{
				var boolValue = (bool)((l is bool) ? l : r);
				var nonBoolValue = (l is bool) ? r : l;
				if (!IsDotNetNumericType(nonBoolValue))
					return false;
				return (boolValue && (Convert.ToDouble(nonBoolValue) == -1)) || (!boolValue && (Convert.ToDouble(nonBoolValue) == 0));
			}

			// Now consider numbers on one or both sides - all special cases are out of the way now so they're either equal or they're not (both
			// sides must be numbers, otherwise it's a non-match)
			if (IsDotNetNumericType(l) && IsDotNetNumericType(r))
				return Convert.ToDouble(l) == Convert.ToDouble(r);
			else if (IsDotNetNumericType(l) || IsDotNetNumericType(r))
				return false;

			// Now do the same for strings and then dates - same deal; they must have consistent types AND values
			if ((l is string) && (r is string))
				return (string)l == (string)r;
			else if ((l is string) || (r is string))
				return false;
			if ((l is DateTime) && (r is DateTime))
				return (DateTime)l == (DateTime)r;

			// Frankly, if we get here then I have no idea what's happened. It will be much easier to identify issues (if any are encountered) if an
			// exception is raised rather than a false response return
			throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
		}

		public object NOTEQ(object l, object r)
		{
			// We can just reverse EQ_Internal's result here, unless it returns null - if it returns null then it means that comparison was not
			// meaningful (one or both sides were DBNull.Value) and so DBNull.Value should be returned.
			var opposingEqualityResult = EQ_Internal(l, r);
			if (opposingEqualityResult == null)
				return null;
			return !opposingEqualityResult.Value;
		}

		public object LT(object l, object r) { return ToVBScriptNullable(LT_Internal(l, r, allowEquals: false)); }
		public object LTE(object l, object r) { return ToVBScriptNullable(LT_Internal(l, r, allowEquals: true)); }
		/// <summary>
		/// This takes the logic from LT but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which LT has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		public bool StrictLT(object l, object r)
		{
			var result = LT_Internal(l, r, allowEquals: false);
			if (result == null)
				throw new InvalidUseOfNullException();
			return result.Value;
		}
		/// <summary>
		/// This takes the logic from LTE but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which LTE has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		public bool StrictLTE(object l, object r)
		{
			var result = LT_Internal(l, r, allowEquals: true);
			if (result == null)
				throw new InvalidUseOfNullException();
			return result.Value;
		}
		private bool? LT_Internal(object l, object r, bool allowEquals)
		{
			// Both sides of the comparison must be simple VBScript values (ie. not object references) - pushing both values through VAL will handle
			// that (an exception will be raised if this operation fails and the value will not be affect if it was already an acceptable type)
			l = _valueRetriever.VAL(l);
			r = _valueRetriever.VAL(r);

			// If one or both sides of the comparison as VBScript Null then that is what is returned
			if ((l == DBNull.Value) || (r == DBNull.Value))
				return null;

			// Check the equality case first, since there may be an early exit we can make (this should return a true or false since the "Null" cases
			// have been handled) - if the values ARE equal then either return true (if allowEquals is true) or false (if allowEquals is false). If
			// not then we'll have to do more work.
			var eq = EQ_Internal(l, r);
			if (eq == null)
				throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
			if (eq.Value)
				return allowEquals;

			// Deal with string special cases next - if both are strings then perform a string comparison. If only one is a string, and it is not blank,
			// then that value is bigger (so if it's on the left then return false and if it's on the right then return true). Blank strings get special
			// handling and are effectively treated as zero (see further down).
			var lString = l as string;
			var rString = r as string;
			if ((lString != null) && (rString != null))
			{
				var stringComparisonResult = STRCOMP_Internal(lString, rString, 0);
				if ((stringComparisonResult == null) || (stringComparisonResult.Value == 0))
					throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
				return stringComparisonResult.Value < 0;
			}
			if ((lString != null) && (lString != ""))
				return false;
			if ((rString != null) && (rString != ""))
				return true;

			// Now we should only have values which can treated as numeric
			// - Actual numbers
			// - Booleans (which return zero or minus one when passed through CDBL)
			// - Null aka VBScript Empty (which returns zero when passed through CDBL)
			// - Blank strings (which can not be passed through CDBL without causing an error, but which we can treat as zero)
			var lNumeric = (lString == "") ? 0 : CDBL_Precise(l);
			var rNumeric = (rString == "") ? 0 : CDBL_Precise(r);
			return lNumeric < rNumeric;
		}

		public object GT(object l, object r) { return ToVBScriptNullable(GT_Internal(l, r, allowEquals: false)); }
		public object GTE(object l, object r) { return ToVBScriptNullable(GT_Internal(l, r, allowEquals: true)); }
		/// <summary>
		/// This takes the logic from GT but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which GT has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		public bool StrictGT(object l, object r)
		{
			var result = GT_Internal(l, r, allowEquals: false);
			if (result == null)
				throw new InvalidUseOfNullException();
			return result.Value;
		}
		/// <summary>
		/// This takes the logic from GTE but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
		/// return a boolean, rather than an object - which GTE has to since it may return a boolean OR DBNull.Value)
		/// </summary>
		public bool StrictGTE(object l, object r)
		{
			var result = GT_Internal(l, r, allowEquals: true);
			if (result == null)
				throw new InvalidUseOfNullException();
			return result.Value;
		}
		private bool? GT_Internal(object l, object r, bool allowEquals)
		{
			// This can just LT_Internal, rather than trying to deal with too much logic itself. When calling LT_Internal, the "allowEquals" value must be
			// the opposite of what we have here - if we are considering GTE then we want !LT (since the equality case should be a match here and not a
			// result which is inverted), if we are considering GT here then we want !LTE (since then equality case would not be a match and LTE would
			// return true for equal l and r values and we would want to invert that result). If LT_Internal returns null, then it means that the
			// comparison is not meaningful (in other words, DBNull.Value was on one or both sides and so DBNull.Value should be returned for
			// any comparison - whether EQ, NOTEQ, LT, GT, etc..)
			var opposingLessThanResult = LT_Internal(l, r, !allowEquals);
			if (opposingLessThanResult == null)
				return null;
			return !opposingLessThanResult.Value;
		}

		public bool IS(object l, object r)
		{
			if (IsVBScriptNothing(l) && IsVBScriptNothing(r))
				return true;
			return _valueRetriever.OBJ(l, "'Is'") == _valueRetriever.OBJ(r, "'Is'");
		}
		public object EQV(object l, object r) { throw new NotImplementedException(); }
		public object IMP(object l, object r) { throw new NotImplementedException(); }

		// Builtin functions - TODO: These are not fully specified yet (eg. LEFT requires more than one parameter and INSTR requires multiple parameters and
		// overloads to deal with optional parameters)
		// - Type conversions
		public byte CBYTE(object value) { return CBYTE(value, "'CByte'"); }
		private byte CBYTE(object value, string exceptionMessageForInvalidContent) { return GetAsNumber<byte>(value, exceptionMessageForInvalidContent, Convert.ToByte); }
		public bool CBOOL(object value) { return _valueRetriever.BOOL(value, "'CBool'"); }
		public decimal CCUR(object value) { return CCUR(value, "'CCur'"); }
		private decimal CCUR(object value, string exceptionMessageForInvalidContent)
		{
			var currencyValue = GetAsNumber<decimal>(value, exceptionMessageForInvalidContent, Convert.ToDecimal);
			if ((currencyValue < VBScriptConstants.MinCurrencyValue) || (currencyValue > VBScriptConstants.MaxCurrencyValue))
				throw new VBScriptOverflowException("'CCur' (" + currencyValue.ToString() + ")");
			return currencyValue;
		}
		public double CDBL(object value)
		{
			// When working with CDBL / CDATE, it seemed like some precision was getting lost when values are passed back and forth through them - eg. if 40000.01
			// is passed into CDATE and then back through CDBL then 40000.01 should come back out. This can be emulate with a double-decimal-double conversion so
			// that these sort of translations seem consistent. However, when trying to parse numbers for other purposes internally, this shouldn't be done since
			// the precision may be important (there are some edge cases in DATEADD where this applies - eg. adding 1.999999999999999 seconds (15x 9s) to a date
			// results in 1 second being added, while adding 1.9999999999999999 seconds (16x 9s) results in 2 seconds being added. I don't think there's a way
			// to perfectly recreate all of VBScript's precision oddities in all cases, so I'm just trying to stick to it being consistent in as many places
			// as possible (which unfortunately means that there's a discrepancy between the internal and public CDBL implementations here).
			return (double)((decimal)CDBL_Precise(value, "'CDbl'"));
		}
		private double CDBL_Precise(object value) { return CDBL_Precise(value, null); }
		private double CDBL_Precise(object value, string optionalExceptionMessageForInvalidContent)
		{
			return GetAsNumber<double>(value, optionalExceptionMessageForInvalidContent, Convert.ToDouble);
		}
		public DateTime CDATE(object value) { return CDATE(value, "'CDate'"); }
		private DateTime CDATE(object value, string exceptionMessageForInvalidContent)
		{
			if (string.IsNullOrWhiteSpace(exceptionMessageForInvalidContent))
				throw new ArgumentException("Null/blank exceptionMessageForInvalidContent specified");

			// Hand off all parsing here to the base valueRetriever.DATE to avoid code duplication
			return _valueRetriever.DATE(value, exceptionMessageForInvalidContent);
		}
		public Int16 CINT(object value) { return CINT(value, "'CInt'"); }
		private Int16 CINT(object value, string exceptionMessageForInvalidContent) { return GetAsNumber<Int16>(value, exceptionMessageForInvalidContent, Convert.ToInt16); }
		public int CLNG(object value) { return CLNG(value, "'CLng'"); }
		public int CLNG(object value, string exceptionMessageForInvalidContent) { return GetAsNumber<int>(value, exceptionMessageForInvalidContent, Convert.ToInt32); }
		public float CSNG(object value) { return CSNG(value, "'CSng'"); }
		private float CSNG(object value, string exceptionMessageForInvalidContent) { return GetAsNumber<float>(value, exceptionMessageForInvalidContent, Convert.ToSingle); }
		public string CSTR(object value) { return CSTR(value, "'CStr'"); }
		private string CSTR(object value, string exceptionMessageForInvalidContent)
		{
			if (string.IsNullOrWhiteSpace(exceptionMessageForInvalidContent))
				throw new ArgumentException("Null/blank exceptionMessageForInvalidContent specified");

			// Hand off all parsing here to the base valueRetriever.STR to avoid code duplication
			return _valueRetriever.STR(value, exceptionMessageForInvalidContent);
		}
		public object INT(object value)
		{
			value = _valueRetriever.VAL(value);

			// Deal with null-like cases
			if (value == DBNull.Value)
				return value;
			if (value == null)
				return (Int16)0;

			// Deal with value type that don't need changing
			if ((value is byte) || (value is Int16) || (value is Int32))
				return value;

			// Deal with a couple of simple case; boolean -> Int16 and Date -> Date (though without any time component)
			if (value is bool)
				return (Int16)((bool)value ? -1 : 0);
			if (value is DateTime)
				return ((DateTime)value).Date;
			var valueWasSingle = value is Single;
			var valueWasDecimal = value is Decimal;
			var valueDouble = GetAsNumber<double>(value, "'Int' (" + value.ToString() + ")", Convert.ToDouble);
			valueDouble = Math.Floor(valueDouble);
			if (valueWasSingle)
				return (Single)valueDouble;
			else if (valueWasDecimal)
				return (Decimal)valueDouble;
			return valueDouble;
		}
		public string STRING(object numberOfTimesToRepeat, object character)
		{
			character = _valueRetriever.VAL(character, "'String'");
			numberOfTimesToRepeat = _valueRetriever.VAL(numberOfTimesToRepeat, "'String'");
			if ((numberOfTimesToRepeat == DBNull.Value) || (character == DBNull.Value))
				throw new InvalidUseOfNullException("'String'");
			int numberOfTimesToRepeatNumber;
			if (numberOfTimesToRepeat == null)
				numberOfTimesToRepeatNumber = 0;
			else
			{
				numberOfTimesToRepeatNumber = CLNG(numberOfTimesToRepeat, "'String'");
				if (numberOfTimesToRepeatNumber < 0)
					throw new InvalidProcedureCallOrArgumentException("'String'");
			}
			char characterChar;
			if (character == null)
				characterChar = '\0';
			else
			{
				var characterString = character as string;
				if (characterString != null)
				{
					if (characterString == "")
						throw new InvalidProcedureCallOrArgumentException("'String'");
					characterChar = characterString[0];
				}
				else
				{
					var characterCode = CINT(character, "'String'");
					if (characterCode > 256)
						characterCode = (short)(characterCode % 256);
					else if (characterCode < 0)
					{
						var numberOf256sToAdd = Math.Ceiling(Math.Abs((double)characterCode / 256));
						characterCode += (short)(numberOf256sToAdd * 256);
					}
					characterChar = (char)characterCode;
				}
			}
			if (numberOfTimesToRepeatNumber > MAX_VBSCRIPT_STRING_LENGTH)
				throw new OutOfStringSpaceException("'String'");
			if (numberOfTimesToRepeatNumber == 0)
				return "";
			return new string(characterChar, numberOfTimesToRepeatNumber);
		}
		// - Number functions
		public object ABS(object value) { throw new NotImplementedException(); }
		public object ATN(object value) { throw new NotImplementedException(); }
		public object COS(object value) { throw new NotImplementedException(); }
		public object EXP(object value) { throw new NotImplementedException(); }
		public object FIX(object value) { throw new NotImplementedException(); }
		public object LOG(object value) { throw new NotImplementedException(); }
		public object OCT(object value) { throw new NotImplementedException(); }
		public object RND(object value) { throw new NotImplementedException(); }
		public object ROUND(object value) { throw new NotImplementedException(); }
		public object SGN(object value) { throw new NotImplementedException(); }
		public object SIN(object value) { throw new NotImplementedException(); }
		public object SQR(object value) { throw new NotImplementedException(); }
		public object TAN(object value) { throw new NotImplementedException(); }
		// - String functions
		public object ASC(object value) { throw new NotImplementedException(); }
		public object ASCB(object value) { throw new NotImplementedException(); }
		public object ASCW(object value) { throw new NotImplementedException(); }
		public string CHR(object value)
		{
			try
			{
				return new string((char)CBYTE(value), 1);
			}
			catch (VBScriptOverflowException e)
			{
				throw new InvalidProcedureCallOrArgumentException("'CHR'", e);
			}
		}
		public object CHRB(object value) { throw new NotImplementedException(); }
		public object CHRW(object value) { throw new NotImplementedException(); }
		public object FILTER(object value) { throw new NotImplementedException(); }
		public object FORMATCURRENCY(object value) { throw new NotImplementedException(); }
		public object FORMATDATETIME(object value) { throw new NotImplementedException(); }
		public object FORMATNUMBER(object value) { throw new NotImplementedException(); }
		public object FORMATPERCENT(object value) { throw new NotImplementedException(); }
		public object HEX(object value) { throw new NotImplementedException(); }

		public object INSTR(object valueToSearch, object valueToSearchFor) { return INSTR(1, valueToSearch, valueToSearchFor); }
		public object INSTR(object startIndex, object valueToSearch, object valueToSearchFor) { return INSTR(startIndex, valueToSearch, valueToSearchFor, 0); }
		public object INSTR(object startIndex, object valueToSearch, object valueToSearchFor, object compareMode)
		{
			// Validate input
			startIndex = _valueRetriever.VAL(startIndex, "'InStr'");
			valueToSearch = _valueRetriever.VAL(valueToSearch, "'InStr'");
			valueToSearchFor = _valueRetriever.VAL(valueToSearchFor, "'InStr'");
			compareMode = _valueRetriever.VAL(compareMode, "'InStr'");
			if (startIndex == DBNull.Value)
				throw new InvalidUseOfNullException("startIndex may not be null");
			var startIndexInt = CLNG(startIndex, "'InStr'");
			if (startIndexInt <= 0)
				throw new InvalidProcedureCallOrArgumentException("'INSTR' (startIndex must be a positive integer)");
			if (compareMode == DBNull.Value)
				throw new InvalidUseOfNullException("compareMode may not be null");
			var compareModeInt = CLNG(compareMode, "'InStr'");
			if ((compareModeInt != 0) && (compareModeInt != 1))
				throw new InvalidProcedureCallOrArgumentException("'INSTR' (compareMode may only be 0 or 1)");

			// Deal with null-ish special cases
			if ((valueToSearch == DBNull.Value) || (valueToSearchFor == DBNull.Value))
				return DBNull.Value;
			if (valueToSearch == null)
				return 0;
			if (valueToSearchFor == null)
				return 1;

			// If the startIndex would go past the end of valueToSearch then return zero
			// - Since startIndex is one-based, we need to subtract one from it to perform this test
			var valueToSearchString = _valueRetriever.STR(valueToSearch);
			var valueToSearchForString = _valueRetriever.STR(valueToSearchFor);
			if (valueToSearchForString.Length + (startIndexInt - 1) > valueToSearchString.Length)
				return 0;

			var useCaseInsensitiveTextComparisonMode = (compareModeInt == 1);
			var zeroBasedMatchIndex = valueToSearchString.IndexOf(
				valueToSearchForString,
				startIndexInt - 1, // This is one-based in VBScript but zero-based in C# (hence the minus one)
				useCaseInsensitiveTextComparisonMode ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal
			);
			return zeroBasedMatchIndex + 1;
		}

		public object INSTRREV(object valueToSearch, object valueToSearchFor)
		{
			// Unlike INSTR, we have to do some work if no startIndex is specified since the default value should be indicate the last character in
			// valueToSearch, if that can be transformed into a non-blank string (if it can not be transformed into a non-object reference at all then
			// throw an exception, and if it is considered to be the equivalent of blank string then default to a startIndex of one, since it's not
			// valid to have a startIndex of zero)
			valueToSearch = _valueRetriever.VAL(valueToSearch, "'InStrRev'");
			int startIndex;
			if ((valueToSearch == null) || (valueToSearch == DBNull.Value))
				startIndex = 1;
			else
				startIndex = Math.Max(1, _valueRetriever.STR(valueToSearch).Length);
			return INSTRREV(valueToSearch, valueToSearchFor, startIndex);
		}
		public object INSTRREV(object valueToSearch, object valueToSearchFor, object startIndex) { return INSTRREV(valueToSearch, valueToSearchFor, startIndex, 0); }
		public object INSTRREV(object valueToSearch, object valueToSearchFor, object startIndex, object compareMode)
		{
			// Validate input
			startIndex = _valueRetriever.VAL(startIndex, "'InStrRev'");
			valueToSearch = _valueRetriever.VAL(valueToSearch, "'InStrRev'");
			valueToSearchFor = _valueRetriever.VAL(valueToSearchFor, "'InStrRev'");
			compareMode = _valueRetriever.VAL(compareMode, "'InStrRev'");
			if (startIndex == DBNull.Value)
				throw new InvalidUseOfNullException("startIndex may not be null");
			var startIndexInt = CLNG(startIndex, "'InStrRev'");
			if (startIndexInt <= 0)
				throw new InvalidProcedureCallOrArgumentException("'INSTRREV' (startIndex must be a positive integer)");
			if (compareMode == DBNull.Value)
				throw new InvalidUseOfNullException("compareMode may not be null");
			var compareModeInt = CLNG(compareMode, "'InStrRev'");
			if ((compareModeInt != 0) && (compareModeInt != 1))
				throw new InvalidProcedureCallOrArgumentException("'INSTRREV' (compareMode may only be 0 or 1)");

			// Deal with null-ish special cases
			if ((valueToSearch == DBNull.Value) || (valueToSearchFor == DBNull.Value))
				return DBNull.Value;
			if (valueToSearch == null)
				return 0;
			if (valueToSearchFor == null)
				return 1;

			// For INSTRREV, the startIndex is taken from the start of the string, like INSTR. But, unlike INSTR, the content to consider is the content
			// preceding this point, rather than the content following it. As such, there is different past-the-end-of-the-content logic to consider and
			// different substring matching logic to apply.
			// - If the startIndex goes beyond the end of the valueToSearch then no match is allowed, similarly if the startIndex indicates a point in
			//   the valueToSearch where there is insufficient content to match valueToSearchFor
			var valueToSearchString = _valueRetriever.STR(valueToSearch);
			var valueToSearchForString = _valueRetriever.STR(valueToSearchFor);
			if ((startIndexInt > valueToSearchString.Length) || (valueToSearchForString.Length > startIndexInt))
				return 0;

			// When searching for a match, only consider the allowed substring of valueToSearch
			var useCaseInsensitiveTextComparisonMode = (compareModeInt == 1);
			var zeroBasedMatchIndex = valueToSearchString.Substring(0, startIndexInt).LastIndexOf(
				valueToSearchForString,
				useCaseInsensitiveTextComparisonMode ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal
			);
			return zeroBasedMatchIndex + 1;
		}

		public object MID(object value) { throw new NotImplementedException(); }
		public object LEN(object value)
		{
			value = _valueRetriever.VAL(value, "'Len'");
			if (value == null)
				return 0;
			else if (value == DBNull.Value)
				return DBNull.Value;
			return _valueRetriever.STR(value).Length;
		}
		public object LENB(object value) { throw new NotImplementedException(); }
		public object LEFT(object value, object maxLength)
		{
			// Validate inputs first
			value = _valueRetriever.VAL(value, "'Left'");
			maxLength = _valueRetriever.VAL(maxLength, "'Left'");
			if (maxLength == DBNull.Value)
				throw new InvalidUseOfNullException();
			var maxLengthInt = CLNG(maxLength, "'Left'");
			if (maxLengthInt < 0)
				throw new InvalidProcedureCallOrArgumentException("'LEFT' (maxLength may not be a negative value)");

			// Deal with special cases
			if (value == null)
				return "";
			if (value == DBNull.Value)
				return DBNull.Value;

			var valueString = _valueRetriever.STR(value);
			maxLengthInt = Math.Min(valueString.Length, maxLengthInt);
			return valueString.Substring(0, maxLengthInt);
		}
		public object LEFTB(object value, object maxLength) { throw new NotImplementedException(); }
		public object RGB(object value) { throw new NotImplementedException(); }
		public object RIGHT(object value, object maxLength)
		{
			// Validate inputs first
			value = _valueRetriever.VAL(value, "'Right'");
			maxLength = _valueRetriever.VAL(maxLength, "'Right'");
			if (maxLength == DBNull.Value)
				throw new InvalidUseOfNullException();
			var maxLengthInt = CLNG(maxLength, "'Right'");
			if (maxLengthInt < 0)
				throw new InvalidProcedureCallOrArgumentException("'LEFT' (maxLength may not be a negative value)");

			// Deal with special cases
			if (value == null)
				return "";
			if (value == DBNull.Value)
				return DBNull.Value;

			var valueString = _valueRetriever.STR(value);
			maxLengthInt = Math.Min(valueString.Length, maxLengthInt);
			return valueString.Substring(valueString.Length - maxLengthInt);
		}
		public object RIGHTB(object value, object maxLength) { throw new NotImplementedException(); }
		public string REPLACE(object value, object toSearchFor, object toReplaceWith) { return REPLACE(value, toSearchFor, toReplaceWith, 1); }
		public string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex) { return REPLACE(value, toSearchFor, toReplaceWith, startIndex, -1); }
		public string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements) { return REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, 0); }
		public string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode)
		{
			// Input validation / type-enforcing
			compareMode = _valueRetriever.VAL(compareMode, "'Replace'");
			if (compareMode == DBNull.Value)
				throw new InvalidUseOfNullException("'Replace'");
			var compareModeNumber = CLNG(compareMode, "'Replace'");
			if ((compareModeNumber != 0) && (compareModeNumber != 1))
				throw new InvalidProcedureCallOrArgumentException("'Replace'");
			maxNumberOfReplacements = _valueRetriever.VAL(maxNumberOfReplacements, "'Replace'");
			if (maxNumberOfReplacements == DBNull.Value)
				throw new InvalidUseOfNullException("'Replace'");
			var maxNumberOfReplacementsNumber = CLNG(maxNumberOfReplacements);
			if (maxNumberOfReplacementsNumber < -1)
				throw new InvalidProcedureCallOrArgumentException("'Replace'");
			startIndex = _valueRetriever.VAL(startIndex, "'Replace'");
			if (startIndex == DBNull.Value)
				throw new InvalidUseOfNullException("'Replace'");
			var startIndexNumber = CLNG(startIndex);
			if ((startIndexNumber < 1) || (startIndexNumber > MAX_VBSCRIPT_STRING_LENGTH))
				throw new InvalidProcedureCallOrArgumentException("'Replace'");
			var toReplaceWithString = _valueRetriever.STR(toReplaceWith, "'Replace'");
			var toSearchForString = _valueRetriever.STR(toSearchFor, "'Replace'");
			var valueString = _valueRetriever.STR(value, "'Replace'");
			if ((maxNumberOfReplacementsNumber == 0) || (valueString == "") || (toSearchForString == "") || (startIndexNumber > valueString.Length)) // Note: VBScript's startIndex is one-based while C#'s is zero-based
				return valueString;

			// Real work
			while ((maxNumberOfReplacementsNumber == -1) || (maxNumberOfReplacementsNumber > 0))
			{
				var replacementIndex = valueString.IndexOf(
					toSearchForString,
					startIndexNumber - 1, // VBScript's startIndex is one-based while C#'s is zero-based
					(compareModeNumber == 0) ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase
				);
				if (replacementIndex == -1)
					break;
				valueString = valueString.Substring(0, replacementIndex) + toReplaceWithString + valueString.Substring(replacementIndex + toSearchForString.Length);
				startIndexNumber = replacementIndex + toReplaceWithString.Length + 1;
				if (maxNumberOfReplacementsNumber != -1)
					maxNumberOfReplacementsNumber--;
			}
			return valueString;
		}
		public object SPACE(object value) { throw new NotImplementedException(); }
		public object[] SPLIT(object value) { return SPLIT(value, " "); }
		public object[] SPLIT(object value, object delimiter)
		{
			// Basic input validation
			delimiter = _valueRetriever.VAL(delimiter, "'Split'");
			if (delimiter == DBNull.Value)
				throw new InvalidUseOfNullException("'Split'");
			value = _valueRetriever.VAL(value, "'Split'");
			if (value == DBNull.Value)
				throw new InvalidUseOfNullException("'Split'");

			// Should be fine to translate both values into strings using the standard mechanism (no exception should arise)
			// Note that Empty and blank string are special cases; always return an empty array, NOT an array with a single element (which would seem more logical)
			// - eg. Split(" ", ",") returns an array with a single element " " while Split("", ",") returns an empty array
			var valueString = _valueRetriever.STR(value, "'Split'");
			var delimiterString = _valueRetriever.STR(delimiter, "'Split'");
			if (string.IsNullOrEmpty(valueString))
				return new object[0];
			return valueString.Split(new[] { delimiterString }, StringSplitOptions.None).Cast<object>().ToArray();
		}
		public object STRCOMP(object string1, object string2) { return STRCOMP(string1, string2, 0); }
		public object STRCOMP(object string1, object string2, object compare) { return ToVBScriptNullable<int>(STRCOMP_Internal(string1, string2, compare)); }
		private int? STRCOMP_Internal(object string1, object string2, object compare)
		{
			throw new NotImplementedException();
		}
		public object STRREVERSE(object value) { throw new NotImplementedException(); }
		public object TRIM(object value)
		{
			value = _valueRetriever.VAL(value, "'Trim'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;
			return _valueRetriever.STR(value).Trim(' ');
		}
		public object LTRIM(object value)
		{
			value = _valueRetriever.VAL(value, "'LTrim'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;
			return _valueRetriever.STR(value).TrimStart(' ');
		}
		public object RTRIM(object value)
		{
			value = _valueRetriever.VAL(value, "'RTrim'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;
			return _valueRetriever.STR(value).TrimEnd(' ');
		}
		public object LCASE(object value)
		{
			value = _valueRetriever.VAL(value, "'LCase'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;
			return _valueRetriever.STR(value).ToLower();
		}
		public object UCASE(object value)
		{
			value = _valueRetriever.VAL(value, "'UCase'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;
			return _valueRetriever.STR(value).ToUpper();
		}
		private const string NonEscapedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@*_+-./";
		public object ESCAPE(object value)
		{
			value = _valueRetriever.VAL(value, "'ESCAPE'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;

			var valueString = _valueRetriever.STR(value);
			if (valueString == "")
				return "";

			var sb = new StringBuilder();
			foreach (var c in valueString)
			{
				if (NonEscapedChars.IndexOf(c) != -1)
				{
					sb.Append(c);
				}
				else if (c <= 0xFF)
				{
					sb.Append("%");
					sb.Append(((int)c).ToString("X2"));
				}
				else
				{
					sb.Append("%u");
					sb.Append(((int)c).ToString("X4"));
				}
			}

			return sb.ToString();
		}
		private static int? HexDigitToInt(char digit)
		{
			if (digit >= '0' && digit <= '9')
				return digit - '0';
			if (digit >= 'A' && digit <= 'F')
				return (digit - 'A') + 0xA;
			if (digit >= 'a' && digit <= 'f')
				return (digit - 'a') + 0xA;

			return null;
		}
		public object UNESCAPE(object value)
		{
			value = _valueRetriever.VAL(value, "'UNESCAPE'");
			if (value == null)
				return "";
			else if (value == DBNull.Value)
				return DBNull.Value;

			var valueString = _valueRetriever.STR(value);
			if (valueString == "")
				return "";

			int length = valueString.Length;
			var sb = new StringBuilder();
			for (int i = 0; i < length; i++)
			{
				if (valueString[i] == '%')
				{
					// Try to parse a %uXXXX sequence
					if (i + 5 < length && valueString[i + 1] == 'u')
					{
						int? digit1 = HexDigitToInt(valueString[i + 2]);
						int? digit2 = HexDigitToInt(valueString[i + 3]);
						int? digit3 = HexDigitToInt(valueString[i + 4]);
						int? digit4 = HexDigitToInt(valueString[i + 5]);

						if (digit1.HasValue && digit2.HasValue && digit3.HasValue && digit4.HasValue)
						{
							sb.Append((char)((digit1 << 12) + (digit2 << 8) + (digit3 << 4) + digit4));
							i += 5;
							continue;
						}
					}

					// Try to parse a %XX sequence
					if (i + 2 < length)
					{
						int? digit1 = HexDigitToInt(valueString[i + 1]);
						int? digit2 = HexDigitToInt(valueString[i + 2]);

						if (digit1.HasValue && digit2.HasValue)
						{
							sb.Append((char)((digit1 << 4) + digit2));
							i += 2;
							continue;
						}
					}
				}

				// Add the character as-is
				sb.Append(valueString[i]);
			}

			return sb.ToString();
		}
		// - Type comparisons
		public bool ISARRAY(object value)
		{
			// Use the same approach as for ISEMPTY..
			try
			{
				bool parameterLessDefaultMemberWasAvailable;
				if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out value))
					return false;
				return (value != null) && value.GetType().IsArray;
			}
			catch (Exception e)
			{
				SETERROR(e);
				return false;
			}
		}
		public bool ISDATE(object value)
		{
			// Use the same basic approach as for ISEMPTY..
			var swallowAnyError = false;
			try
			{
				bool parameterLessDefaultMemberWasAvailable;
				if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out value))
					return false;

				// Any error encountered in evaluating the default member (if required to coerce value into a value type) should be recorded with
				// SETERROR, but if the value is not a valid date and an exception is thrown by the DateParser, then that should NOT be recorded
				swallowAnyError = true;
				if (value == null)
					return false;
				if (value is DateTime)
					return true;
				DateParser.Default.Parse(value.ToString()); // If this doesn't throw an exception then it must be a valid-for-VBScript date string
				return true;
			}
			catch (Exception e)
			{
				if (!swallowAnyError)
					SETERROR(e);
				return false;
			}
		}
		public bool ISEMPTY(object value)
		{
			try
			{
				// If this can not be coerced into a value type then it can't be Empty, so return false
				bool parameterLessDefaultMemberWasAvailable;
				if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out value))
					return false;

				// If it IS a value type, or was manipulated into one, then check for null (aka VBScript's Empty)
				return value == null;
			}
			catch (Exception e)
			{
				// If an exception was raised while evaluating a default member (meaning "value" was not a value but it had a default member that could
				// be investigated.. but an exception was raised within the evaluation of that member) then record the error and return false (as is
				// consistent with VBScript's behaviour)
				SETERROR(e);
				return false;
			}
		}
		public bool ISNULL(object value)
		{
			// Use the same approach as for ISEMPTY..
			try
			{
				bool parameterLessDefaultMemberWasAvailable;
				if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out value))
					return false;
				return value == DBNull.Value;
			}
			catch (Exception e)
			{
				SETERROR(e);
				return false;
			}
		}
		private static Regex SpaceFollowingMinusSignRemover = new Regex(@"-\s+", RegexOptions.Compiled);
		public bool ISNUMERIC(object value)
		{
			// Use the same basic approach as for ISEMPTY..
			try
			{
				bool parameterLessDefaultMemberWasAvailable;
				if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out value))
					return false;
				if (value == null)
					return true; // Empty is identified as numeric in VBScript
				// double.TryParse seems to match VBScript's pretty well (see the test cases for more details) with one exception; VBScript will tolerate whitespace between
				// a negative sign and the start of the content, so we need to do consider replacements (any "-" followed by whitespace should become just "-")
				double numericValue;
				return double.TryParse(SpaceFollowingMinusSignRemover.Replace(value.ToString(), "-"), out numericValue);
			}
			catch (Exception e)
			{
				SETERROR(e);
				return false;
			}
		}
		public bool ISOBJECT(object value)
		{
			return !_valueRetriever.IsVBScriptValueType(value);
		}
		public string TYPENAME(object value)
		{
			if (value == null)
				return "Empty";
			if (value == DBNull.Value)
				return "Null";
			if (IsVBScriptNothing(value))
				return "Nothing";

			var type = value.GetType();
			if (type.IsArray && (type.GetElementType() == typeof(Object)))
				return "Variant()";
			if (_valueRetriever.IsVBScriptValueType(value))
			{
				if (type == typeof(bool))
					return "Boolean";
				if (type == typeof(byte))
					return "Byte";
				if (type == typeof(Int16))
					return "Integer";
				if (type == typeof(Int32))
					return "Long";
				if (type == typeof(double))
					return "Double";
				if (type == typeof(DateTime))
					return "Date";
				if (type == typeof(Decimal))
					return "Currency";
				return Information.TypeName(value);
			}

			if (type.IsCOMObject)
			{
				var typeDescriptorClassName = TypeDescriptor.GetClassName(value);
				if (!string.IsNullOrWhiteSpace(typeDescriptorClassName))
					return typeDescriptorClassName;
			}
			var sourceClassName = type.GetCustomAttributes(typeof(SourceClassName), inherit: true).FirstOrDefault() as SourceClassName;
			if (sourceClassName != null)
				return sourceClassName.Name;

			// This will always fall through to Object if it finds nothing better along the way
			while (true)
			{
				var comVisibleAttributeIfAny = type.GetCustomAttributes(typeof(ComVisibleAttribute), inherit: false).Cast<ComVisibleAttribute>().FirstOrDefault();
				if ((comVisibleAttributeIfAny != null) && comVisibleAttributeIfAny.Value)
					return type.Name;
				type = type.BaseType;
			}
		}
		public object VARTYPE(object value) { throw new NotImplementedException(); }
		// - Array functions
		public object ARRAY(params object[] values)
		{
			if (values == null)
				throw new ArgumentNullException("values");
			return values;
		}
		public void ERASE(object target, Action<object> targetSetter)
		{
			// ERASE is more like a keyword in VBScript than a function - none of the builtin VBScript functions take arguments by-ref and nearly all of them apply a lot of
			// similar handling to inputs such as raising invalid-use-of-null errors where VBScript Null is not expected and considering parameter-less default properties
			// and function when expected a value type and receiving an object reference. ERASE does not do that; if the target is not an array then it's a type mismatch,
			// doesn't matter whether it's Empty, Null, Nothing, a number, a string, a date, an object reference with a default parameterless property; it's type mismatch!
			// - Note: A "targetSetter" is required to update the array, rather than just taking the target argument as by-ref, since it would be common for translated
			//   code to call "_.ERASE(ref outer.names)", which would be invalid C# code since "ref" cannot be used with property accessors
			if ((target == null) || !target.GetType().IsArray)
				throw new TypeMismatchException("'Erase'");
			targetSetter(new object[0]);
		}
		public void ERASE(object target, params object[] arguments)
		{
			// This variation of ERASE is similarly strict to the one above (target must be an array or it's a type mismatch, no matter what!) but the arguments are then
			// evaluated and interpreted as array index values - if this fails then it's a type mismatch as well. The indices must point at an element in the array that
			// is also an array, that is what will get erased. If the argument count does not match the array rank then it's a subscript-out-of-range failure (this
			// includes the case of zero arguments, which is what "ERASE a()" is translated into - it needs to get to this point at runtime so that the type of
			// "a" can be checked, which determines whether the failure is a type-mismatch or subscript-out-of-range).
			var targetArray = target as Array;
			if (targetArray == null)
				throw new TypeMismatchException("'Erase'");
			if ((arguments == null) || (arguments.Length == 0))
				throw new SubscriptOutOfRangeException("'Erase'");
			var numericArguments = arguments.Select(a => CLNG(a, "'Erase'")).ToArray();
			if (targetArray.Rank != numericArguments.Length)
				throw new SubscriptOutOfRangeException("'Erase'");
			object elementValue;
			try
			{
				elementValue = targetArray.GetValue(numericArguments);
			}
			catch (Exception e)
			{
				throw new SubscriptOutOfRangeException("'Erase'", e);
			}
			if ((elementValue as Array) == null)
			{
				// The element in the target array must also so be an array since that is what's effectively getting erased
				throw new TypeMismatchException("'Erase'");
			}
			targetArray.SetValue(new object[0], numericArguments);
		}
		public string JOIN(object value) { return JOIN(value, " "); }
		public string JOIN(object value, object delimiter)
		{
			delimiter = _valueRetriever.VAL(delimiter, "'Join'");
			if (value == DBNull.Value)
				throw new InvalidUseOfNullException("'Join'");
			value = _valueRetriever.VAL(value, "'Join'");
			if (delimiter == DBNull.Value)
				throw new InvalidUseOfNullException("'Join'");
			if (value == null)
				throw new TypeMismatchException("'Join'");
			if (value == DBNull.Value)
				throw new InvalidUseOfNullException("'Join'");
			var valueType = value.GetType();
			if (!valueType.IsArray)
				throw new TypeMismatchException("'Join'");
			var arrayRank = valueType.GetArrayRank();
			if (arrayRank == 0)
				return "";
			else if (arrayRank > 1)
				throw new TypeMismatchException("'Join'");
			return string.Join(
				(delimiter == null) ? "" : _valueRetriever.STR(delimiter),
				((Array)value)
					.Cast<object>()
					.Select(element =>
					{
						element = _valueRetriever.VAL(element, "'Join'");
						if (element == DBNull.Value)
							throw new TypeMismatchException("'Join'");
						return (element == null) ? "" : _valueRetriever.STR(element);
					})
			);
		}
		public int LBOUND(object value) { return LBOUND(value, 1); }
		public int LBOUND(object value, object dimension)
		{
			// If both the value and dimension are invalid values, the dimension errors should be raised first (so try to process that value first)
			var dimensionInt = CLNG(dimension, "'LBound'");
			var array = _valueRetriever.VAL(value, "'LBound'") as Array;
			if (array == null)
				throw new TypeMismatchException("'LBound'");
			if ((dimensionInt < 1) || (dimensionInt > array.Rank))
				throw new SubscriptOutOfRangeException("'LBound'");
			return array.GetLowerBound(dimensionInt - 1); // Note: VBScript uses one-based a dimension value here while C# is zero-based, hence the -1
		}
		public int UBOUND(object value) { return UBOUND(value, 1); }
		public int UBOUND(object value, object dimension)
		{
			// If both the value and dimension are invalid values, the dimension errors should be raised first (so try to process that value first)
			var dimensionInt = CLNG(dimension, "'UBound'");
			var array = _valueRetriever.VAL(value, "'UBound'") as Array;
			if (array == null)
				throw new TypeMismatchException("'UBound'");
			if ((dimensionInt < 1) || (dimensionInt > array.Rank))
				throw new SubscriptOutOfRangeException("'UBound'");
			return array.GetUpperBound(dimensionInt - 1); // Note: VBScript uses one-based a dimension value here while C# is zero-based, hence the -1
		}
		// - Date functions
		public DateTime NOW() { return DateTime.Now; }
		public DateTime DATE() { return DateTime.Now.Date; }
		public DateTime TIME() { return new DateTime(DateTime.Now.TimeOfDay.Ticks); }
		public object DATEADD(object interval, object number, object value)
		{
			// DateAdd seems to be an usual functions - it ignores fractions in "number" rather than rounding them (so adding 101, 101.5 or 101.9 is the same as adding 101).
			// It's also unusual in that it won't overflow for enormous numeric values, it always falls back to an invalid-procedure-call-or-argument error (if the number
			// would result in an unrepresentable date). On top of this, it doesn't validate all of its arguments before considering any work - DateAdd("x", "y", Null)
			// returns Null, despite the fact that the "interval" and "number" arguments are nonsense; DateAdd("x", "y", Now()) would result in a type-mismatch error.
			value = _valueRetriever.VAL(value, "'DateAdd'");
			if (value == DBNull.Value)
				return DBNull.Value; // Don't even check the other arguments if we got a Null value argument
			var dateValue = CDATE(value, "'DateAdd'");
			// The MSDN documentation (for VBA, but which is the closest I could find: https://msdn.microsoft.com/en-us/library/aa262710%28v=vs.60%29.aspx) says that "If
			// number isn't a Long value, it is rounded to the nearest whole number before being evaluated." However, testing with VBScript shows this not to be the case.
			// For example, adding (for any interval) 103, 103.1, 103.5 or 103.9 all have the same effect, as do adding 102, 102.1, 102.5 or 102.9, which indicates that
			// the fractional part of the number is being ignored, not rounded. Pushing the limits shows that 1.999999999999999 (15x 9s) will result in 1 being added while
			// 1.9999999999999999 (16x 9s) will result in 2 being added. With 10.9999 it's still 15 vs 16 9s where it changes (from 10 to 11), while with 100.999 it's
			// 14 vs 15 9s. This is consistent with double precision in .net and using CDBL_Precise and then truncating the value will achieve the same effect.
			// - On top of this, if the number lies outside the Int32 range ("Long" in VBScript), then it initially looks like it rolls over.. but actually it just rolls
			//   over and gets stuck at Int32.MinValue; for example any of the following number values will result in the same as if -2147483648 (Int32.MinValue) had been
			//   specified as the number argument: 2147483648 (Int32.MaxValue + 1), 21474836470 (Int32.MaxValue * 10), 1844674407370955161500 (UInt64.MaxValue * 10)
			int intNumber;
			var doubleNumber = Math.Truncate(CDBL_Precise(number, "'DateAdd'"));
			if ((doubleNumber < int.MinValue) || (doubleNumber > int.MaxValue))
				intNumber = int.MinValue;
			else
				intNumber = (int)doubleNumber;
			interval = _valueRetriever.VAL(interval, "'DateAdd'");
			if (interval == DBNull.Value)
				throw new InvalidUseOfNullException("'DateAdd'");
			var intervalString = interval as string;
			if (intervalString == null)
				throw new InvalidProcedureCallOrArgumentException("'DateAdd'");
			Func<DateTime, int, DateTime> dateManipulator;
			switch (intervalString.ToLower()) // Interval matching is case-insensitive in VBScript (it won't allow leading or trailing whitespace, though)
			{
				default:
					throw new InvalidProcedureCallOrArgumentException("'DateAdd'");
				case "yyyy":
					dateManipulator = (date, increment) => date.AddYears(increment);
					break;
				case "q":
					dateManipulator = (date, increment) => date.AddMonths(increment * 3); // quarter
					break;
				case "m":
					dateManipulator = (date, increment) => date.AddMonths(increment);
					break;
				case "ww":
					dateManipulator = (date, increment) => date.AddDays(increment * 7); // week
					break;
				case "y":
				case "d":
				case "w":
					// Any of "y" (Day of year), "d" (Day) or "w" (weekday) may be used to alter the date, apparently, according to an MSDN article (but this also says that fractional number
					// values are rounded to the NEAREST whole number, which they aren't, so what does it know.. https://msdn.microsoft.com/en-us/library/aa262710%28v=vs.60%29.aspx). Presumably
					// these three values are all supported for consistency with related functions such as DATEPART, where the three values will NOT act the same)
					dateManipulator = (date, increment) => date.AddDays(increment);
					break;
				case "h":
					dateManipulator = (date, increment) => date.AddHours(increment);
					break;
				case "n":
					dateManipulator = (date, increment) => date.AddMinutes(increment); // This is minutes since "m" is used for months (and don't differentiate between "M" and "m", unlike .net)
					break;
				case "s":
					dateManipulator = (date, increment) => date.AddSeconds(increment);
					break;
			}
			try
			{
				dateValue = dateManipulator(dateValue, intNumber);
			}
			catch (Exception e)
			{
				throw new InvalidProcedureCallOrArgumentException("'DateAdd'", e);
			}
			if ((dateValue < VBScriptConstants.EarliestPossibleDate) || (dateValue.Date > VBScriptConstants.LatestPossibleDate.Date))
				throw new InvalidProcedureCallOrArgumentException("'DateAdd'");
			return dateValue;
		}
		public object DATEDIFF(object value) { throw new NotImplementedException(); }
		public object DATEPART(object value) { throw new NotImplementedException(); }
		public object DATESERIAL(object year, object month, object date)
		{
			// TODO: This is not a complete implementation, it's just enough to get moving

			// TODO: Implement (and write tests) for this more thoroughly - eg. (99,2,10) => 1999-2-10, (99,14,10) => 100-2-10, (2017,13,1) => 2018-1-1

			var numericYear = CLNG(year);
			var numericMonth = CLNG(month);
			var numericDate = CLNG(date);

			if ((numericMonth < 0) || (numericMonth > 12))
			{
				var numberOfYearsToAdd = (int)Math.Floor((double)numericMonth / 12);
				numericYear += numberOfYearsToAdd;
				numericMonth = numericMonth % 12;
				if (numericMonth < 0)
					numericMonth += 12; // For negative values (eg. -1 % 12 is -1 so need to add 12 to get to 11, -13 % 12 is also -1 so never need to add more or less than 12)
			}

			// TODO: Check days <= 0 or days > days-in-month/year

			// TODO: Check small year values
			// TODO: What about negative year values??

			return new DateTime(numericYear, numericMonth, numericDate);
		}
		public DateTime DATEVALUE(object value)
		{
			// In summary, this will do a subset of the processing of CDATE (it will accept a DateTime or a parse-able string, but not a numeric value such as 123.45) and return only the date:
			//   "The reasons for using DateValue and TimeValue to convert a string instead of CDate may not be immediately obvious. Consider the example above. CDate is creating a Date value for the entire supplied
			//    string.  DateValue and TimeValue will allow you to create Date values containing only the specified portion of the string while ignoring the rest."
			// - http://www.aspfree.com/c/a/windows-scripting/working-with-dates-and-times-in-vbscript/
			value = _valueRetriever.VAL(value, "'DateValue'");
			if (value == null)
				throw new TypeMismatchException("'DateValue'");
			if (value == DBNull.Value)
				throw new InvalidUseOfNullException("'DateValue'");
			DateTime dateValue;
			if (value is DateTime)
				dateValue = (DateTime)value;
			else
			{
				try
				{
					dateValue = DateParser.Default.Parse(value.ToString());
				}
				catch (Exception e)
				{
					throw new TypeMismatchException("'DateValue'", e);
				}
			}
			return dateValue.Date;
		}
		public object TIMESERIAL(object value) { throw new NotImplementedException(); }
		public DateTime TIMEVALUE(object value)
		{
			// In summary, this will do a subset of the processing of CDATE (it will accept a DateTime or a parse-able string, but not a numeric value such as 123.45) and return only the time component:
			//   "The reasons for using DateValue and TimeValue to convert a string instead of CDate may not be immediately obvious. Consider the example above. CDate is creating a Date value for the entire supplied
			//    string.  DateValue and TimeValue will allow you to create Date values containing only the specified portion of the string while ignoring the rest."
			// - http://www.aspfree.com/c/a/windows-scripting/working-with-dates-and-times-in-vbscript/
			value = _valueRetriever.VAL(value, "'TimeValue'");
			if (value == null)
				throw new TypeMismatchException("'TimeValue'");
			if (value == DBNull.Value)
				throw new InvalidUseOfNullException("'TimeValue'");
			DateTime dateValue;
			if (value is DateTime)
				dateValue = (DateTime)value;
			else
			{
				try
				{
					dateValue = DateParser.Default.Parse(value.ToString());
				}
				catch (Exception e)
				{
					throw new TypeMismatchException("'TimeValue'", e);
				}
			}
			// VBScript represents times by taking its "zero date" and adding hours / minutes / seconds to it
			return VBScriptConstants.ZeroDate.Add(dateValue.TimeOfDay);
		}
		public object DAY(object value)
		{
			value = _valueRetriever.VAL(value, "'Day'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			return ToClosestSecond(CDATE(value, "'Day'")).Day;
		}
		public object MONTH(object value)
		{
			value = _valueRetriever.VAL(value, "'Month'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			return ToClosestSecond(CDATE(value, "'Month'")).Month;
		}
		public object MONTHNAME(object value) { throw new NotImplementedException(); }
		public object YEAR(object value)
		{
			value = _valueRetriever.VAL(value, "'Year'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			return ToClosestSecond(CDATE(value, "'Year'")).Year;
		}
		public object WEEKDAY(object value)
		{
			return WEEKDAY(value, VBScriptConstants.vbSunday);
		}
		public object WEEKDAY(object value, object firstDayOfWeek)
		{
			value = _valueRetriever.VAL(value, "'Weekday'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			var date = ToClosestSecond(CDATE(value, "'Weekday'"));

			// NOTE: VBScript weekdays go from Sunday (1) to Saturday (7) (unless overriden by firstDayOfWeek), while .NET DayOfWeek goes from Sunday (0) to Saturday (6)
			var vbsFirstDayOfWeek = CLNG(firstDayOfWeek, "'Weekday'");
			if (vbsFirstDayOfWeek < 0 || vbsFirstDayOfWeek > 7)
				throw new InvalidProcedureCallOrArgumentException("'Weekday'");
			if (vbsFirstDayOfWeek == VBScriptConstants.vbUseSystemDayOfWeek)
				vbsFirstDayOfWeek = (int)CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek + 1;

			return (((int)date.DayOfWeek + (8 - vbsFirstDayOfWeek)) % 7) + 1;
		}
		public object WEEKDAYNAME(object value) { throw new NotImplementedException(); }
		public object HOUR(object value)
		{
			value = _valueRetriever.VAL(value, "'Hour'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			return ToClosestSecond(CDATE(value, "'Hour'")).Hour;
		}
		public object MINUTE(object value)
		{
			value = _valueRetriever.VAL(value, "'Minute'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			return ToClosestSecond(CDATE(value, "'Minute'")).Minute;
		}
		public object SECOND(object value)
		{
			value = _valueRetriever.VAL(value, "'Second'");
			if (value == DBNull.Value)
				return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
			return ToClosestSecond(CDATE(value, "'Second'")).Second;
		}
		private DateTime ToClosestSecond(DateTime value)
		{
			var approximateValue = new DateTime(value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second);
			if (value.Millisecond >= 500) // TODO: Check whether this rounding is correct, should it be banker's rounding?
				approximateValue = approximateValue.AddSeconds(1);
			return approximateValue;
		}
		// - Object creation
		public object CREATEOBJECT(object value) { throw new NotImplementedException(); }
		public object GETOBJECT(object value) { throw new NotImplementedException(); }
		public object EVAL(object value) { throw new NotImplementedException(); }
		public object EXECUTE(object value) { throw new NotImplementedException(); }
		public object EXECUTEGLOBAL(object value) { throw new NotImplementedException(); }
		// - Misc
		public object GETLOCALE(object value) { throw new NotImplementedException(); }
		public object GETREF(object value) { throw new NotImplementedException(); }
		public object INPUTBOX(object value) { throw new NotImplementedException(); }
		public object LOADPICTURE(object value) { throw new NotImplementedException(); }
		public object MSGBOX(object value) { throw new NotImplementedException(); }
		public string SCRIPTENGINE(object value) { throw new NotImplementedException(); }
		public int SCRIPTENGINEBUILDVERSION(object value) { throw new NotImplementedException(); }
		public int SCRIPTENGINEMAJORVERSION(object value) { throw new NotImplementedException(); }
		public int SCRIPTENGINEMINORVERSION(object value) { throw new NotImplementedException(); }
		public object SETLOCALE(object value) { throw new NotImplementedException(); }

		/// <summary>
		/// This returns the value without any immediate processing, but may keep a reference to it and dispose of it (where applicable) after
		/// the request completes (to try to avoid resources from not being cleaned up in the absence of the VBScript deterministic garbage
		/// collection - classes with a Class_Terminate function are translated into IDisposable types and, while IDisposable.Dispose will not
		/// be called by the translated code, it may be called after the request ends if the requests are tracked here. This will throw an
		/// exception for a null value.
		/// </summary>
		public object NEW(object value)
		{
			if (value == null)
				throw new ArgumentNullException("value");

			var disposableResource = value as IDisposable;
			if (disposableResource != null)
				_disposableReferencesToClearAfterTheRequest.Add(disposableResource);
			return value;
		}

		// Array definitions
		public object NEWARRAY(IEnumerable<object> dimensions)
		{
			if (dimensions == null)
				throw new ArgumentNullException("dimensions");

			// Note that VBScript specifies upper bounds for arrays, rather than the size - so ReDim a(2) means that the array "a" needs three
			// elements (0, 1 and 2) and so must be declared in C# as object[3]. In VBScript, if negative ranges are specified below -1 (since
			// -1 means zero in C#, which is not an unreasonable request - eg. object[0]) then an out-of-memory error is raised. It shouldn't
			// be possible for this to be called without any dimensions from translated code since that would be a syntax error (and so may
			// be an ArgumentException rather than a specialise VBScript exception).
			var dimensionSizes = dimensions.Select(d => CLNG(d, "'NewArray'") + 1).ToArray();
			if (!dimensionSizes.Any())
				throw new ArgumentException("No dimensions specified for NEWARRAY");
			if (dimensionSizes.Any(d => d < 0))
				throw new OutOfMemoryException("Invalid negative dimensions used for NEWARRAY call");
			return Array.CreateInstance(typeof(object), dimensionSizes);
		}

		public object RESIZEARRAY(object array, IEnumerable<object> dimensions)
		{
			// Note: Don't even check "array" for null until the dimensions have been evaluated
			if (dimensions == null)
				throw new ArgumentNullException("dimensions");

			// Note that VBScript specifies upper bounds for arrays, rather than the size - so ReDim a(2) means that the array "a" needs three
			// elements (0, 1 and 2) and so must be declared in C# as object[3]. In VBScript, if negative ranges are specified below -1 (since
			// -1 means zero in C#, which is not an unreasonable request - eg. object[0]) then an out-of-memory error is raised. It shouldn't
			// be possible for this to be called without any dimensions from translated code since that would be a syntax error (and so may
			// be an ArgumentException rather than a specialise VBScript exception).
			// - The dimensions are evaulated before the target array is validated (before it is even checked for null, even) in order to
			//   be consistent with VBScript's runtime behaviour
			var dimensionSizes = dimensions.Select(d => CLNG(d, "'ResizeArray'") + 1).ToArray();
			if (!dimensionSizes.Any())
				throw new ArgumentException("No dimensions specified for RESIZEARRAY");
			var arrayTyped = array as Array;
			if (arrayTyped == null)
				throw new TypeMismatchException("'ResizeArray' target not an array");
			if (dimensionSizes.Length != arrayTyped.Rank)
				throw new SubscriptOutOfRangeException("Inconsistent number of dimensions specified for RESIZEARRAY");
			if (dimensionSizes.Any(d => d < 0))
				throw new OutOfMemoryException("Invalid negative dimensions used for RESIZEARRAY call");

			for (var dimension = 0; dimension < arrayTyped.Rank - 1; dimension++)
			{
				if (arrayTyped.GetLength(dimension) != dimensionSizes[dimension])
					throw new SubscriptOutOfRangeException("Invalid dimensions specified for RESIZEARRAY - only the last dimension may vary in size");
			}

			if (dimensionSizes.Length == 1)
			{
				// Copying a 1D array is easy..
				var newArray = new object[dimensionSizes[0]];
				Array.Copy(arrayTyped, newArray, Math.Min(arrayTyped.Length, dimensionSizes[0]));
				return newArray;
			}
			else if (dimensionSizes.Length == 2)
			{
				// Copying a 2D array can be done column-by-column, so there's only one loop and an Array.Copy per iteration..
				var newArray = new object[dimensionSizes[0], dimensionSizes[1]];
				var numberOfElementsToCopyEachTime = Math.Min(arrayTyped.GetLength(1), dimensionSizes[1]);
				if (numberOfElementsToCopyEachTime > 0)
				{
					for (var i = 0; i < dimensionSizes[0]; i++)
					{
						Array.Copy(
							arrayTyped,
							i * arrayTyped.GetLength(1),
							newArray,
							i * dimensionSizes[1],
							numberOfElementsToCopyEachTime
						);
					}
				}
				return newArray;
			}
			else
			{
				// Copying an array with more dimensions is more awkward.. the only way I can think of is to go through every element of the
				// new array and copy each value from the old array, so long as the element exists in the old array. This is MUCH less
				// efficient than the process for the 1D or 2D arrays.
				var newArray = Array.CreateInstance(typeof(object), dimensionSizes);
				var totalNumberOfElements = dimensionSizes.Aggregate(1, (acc, value) => acc * value);
				var indicesOfElementToCopy = new int[dimensionSizes.Length];
				for (var i = 0; i < totalNumberOfElements; i++)
				{
					if (i > 0)
					{
						var indexToIncrementNext = 0;
						while (true)
						{
							if (indicesOfElementToCopy[indexToIncrementNext] < (dimensionSizes[indexToIncrementNext] - 1))
							{
								indicesOfElementToCopy[indexToIncrementNext]++;
								break;
							}
							indicesOfElementToCopy[indexToIncrementNext] = 0;
							indexToIncrementNext++;
						}
					}
					var elementDoesNotExistInSource = false;
					for (var j = 0; j < indicesOfElementToCopy.Length; j++)
					{
						if (arrayTyped.GetLength(j) <= indicesOfElementToCopy[j])
						{
							elementDoesNotExistInSource = true;
							break;
						}
					}
					if (!elementDoesNotExistInSource)
						newArray.SetValue(arrayTyped.GetValue(indicesOfElementToCopy), indicesOfElementToCopy);
				}
				return newArray;
			}
		}

		private IEnumerable<int> GetDimensions(IEnumerable<object> dimensions)
		{
			if (dimensions == null)
				throw new ArgumentNullException("dimensions");

			throw new NotImplementedException(); // TODO
		}

		public object NEWREGEXP()
		{
			// TODO: Ideally, the object returned here would be a managed implementation of "VBScript.RegExp" (which has a fairly simple interface), to reduce the
			// number of dependencies. But this works and so will do for the time being.
			return Activator.CreateInstance(Type.GetTypeFromProgID("VBScript.RegExp", throwOnError: true));
		}

		/// <summary>
		/// This will never be null (if there is no error then an ErrorDetails with Number zero will be returned)
		/// </summary>
		public ErrorDetails ERR
		{
			get
			{
				var currentError = _trappedErrorIfAny;
				if (currentError == null)
					return ErrorDetails.NoError;
				var currentErrorAsVBScriptSpecificError = currentError as SpecificVBScriptException;
				return new ErrorDetails(
					(currentErrorAsVBScriptSpecificError != null) ? currentErrorAsVBScriptSpecificError.ErrorNumber : 1, // TODO: Still need a better way to get error number for non-VBScript-specific errors
					currentError.Source,
					currentError.Message,
					originalExceptionIfKnown: currentError
				);
			}
		}

		/// <summary>
		/// There are some occassions when the translated code needs to throw a runtime exception based on the content of the source code - eg.
		///   WScript.Echo 1()
		/// It is clear that "1" is a numeric constant and not a function, and so may not be called as one. However, this is not invalid VBScript and so is
		/// not a compile time error, it is something that must result in an exception at runtime. In these cases, where it is known at the time of translation
		/// that an exception must be thrown, this method may be used to do so at runtime. This is different to SETERROR, since that records an exception that
		/// has already been thrown - this throws the specified exception (it returns an object, rather than void, for the same reason as the below signatures).
		/// </summary>
		public object RAISEERROR(Exception e)
		{
			if (e == null)
				throw new ArgumentNullException("e");

			throw e;
		}

		// These method signatures have to return a value since these are what are called when the source code includes "Err.Raise 123", which VBScript allows
		// to exist in the form "If (Err.Raise(123)) Then" - if these didn't return values then there could be compile errors in the generated C# that were
		// valid VBScript.
		public object RAISEERROR(object number) { return RAISEERROR(number, ""); }
		public object RAISEERROR(object number, object source) { return RAISEERROR(number, source, ""); }
		public object RAISEERROR(object number, object source, object description)
		{
			// This is another function (like ERASE) that doesn't give many clues - almost every failure is a "Type mismatch" (Null values do not result in
			// "Invalid use of null" and Nothing does not result in "Object variable not set"). However, if "number" is zero then the other two arguments
			// are not evaluated - this only happens if the value for number is ok. And if number is zero then it DOES get a different error :S
			int numericNumber;
			try
			{
				numericNumber = CLNG(number);
			}
			catch (Exception e)
			{
				throw new TypeMismatchException("Err.Raise", e);
			}
			if (numericNumber == 0)
				throw new InvalidProcedureCallOrArgumentException("Err.Raise");
			string sourceString, descriptionString;
			try
			{
				sourceString = _valueRetriever.STR(source);
				descriptionString = _valueRetriever.STR(description);
			}
			catch (Exception e)
			{
				throw new TypeMismatchException("Err.Raise", e);
			}
			throw new CustomException(numericNumber, sourceString, descriptionString);
		}

		public void SETERROR(Exception e)
		{
			// Note that there is (at most) only a single error associated with an executing request. If the error-trapping is enabled and a function F1()
			// executes code that raises an error but then goes and calls F2() which also raises an error, the error recorded from the code in F1 that
			// occured before calling F2 is lost, it is overwritten by F2. So there is no need to try to map trapped errors onto error tokens, there is
			// only one per request (or zero - if there has been no error trapped, or if there HAS been an error trapped that has then been cleared).
			if (e == null)
				throw new ArgumentNullException("e");
			_trappedErrorIfAny = e;
		}

		public void CLEARANYERROR()
		{
			// This should be called by translated code that originates from an ON ERROR GOTO 0 with no corresponding ON ERROR RESUME NEXT - the translation
			// process will not emit code to call GETERRORTRAPPINGTOKEN since the source is not trying to trap any errors. However, any error information
			// must be cleared nonetheless, since there was an ON ERROR GOTO 0 in the source. It will also be required when Err.Clear is called.
			_trappedErrorIfAny = null;
		}

		public int GETERRORTRAPPINGTOKEN()
		{
			// Every time error-trapping is enabled within a function (or the outermost scope, where code doesn't run within a function in VBScript), the
			// translated code must request an "error trapping token". This is used to keep track of where error-trapping is and isn't enabled. If, for
			// example, a function F1 includes an ON ERROR RESUME NEXT and then calls F2 which includes its own ON ERROR RESUME NEXT and then later an
			// ON ERROR GOTO 0, this must only disable error-trapping within F2, the error-trapping that was enabled in F1 must continue to be enabled.
			// It isn't known at translation time how many error tokens may be required since this depends upon how the code executes - if F2 calls
			// itself then within its ON ERROR RESUME NEXT .. ON ERROR GOTO 0 region, an ON ERROR GOTO 0 call from that second call to F2 must not
			// disable error-trapping in the context of the first F2 call. So error tokens need to be handled dynamically. To try to only maintain as
			// many as strictly necessary, there is a queue of available tokens that is used to service GETERRORTRAPPINGTOKEN calls - after an error
			// token is returned (through RELEASEERRORTRAPPINGTOKEN), it goes back into the queue to potentially be used again. If the queue is empty
			// when this method is called then a new token is created. The token values are incremented each time this happens to ensure that they are
			// unique. This is why it's important that tokens are properly released - either when error-trapping is disabled (through an explicit ON
			// ERROR GOTO 0 or through an error being trapped or through a function scope ending where ON ERROR RESUME NEXT was set).
			// Note: When tokens are first requested, they default to the "OnErrorGoto0" state - meaning that error-trapping is not enabled currently
			// for that token. Error-trapping is enabled through a subsequent call to STARTERRORTRAPPINGANDCLEARANYERROR.
			int token;
			if (_availableErrorTokens.Any())
				token = _availableErrorTokens.Dequeue();
			else
				token = _availableErrorTokens.Count + _activeErrorTokens.Count + 1;
			_activeErrorTokens.Add(token, ErrorTokenState.OnErrorGoto0);
			return token;
		}

		public void RELEASEERRORTRAPPINGTOKEN(int errorToken)
		{
			if (!_activeErrorTokens.ContainsKey(errorToken))
				throw new Exception("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");
			_activeErrorTokens.Remove(errorToken);
			_availableErrorTokens.Enqueue(errorToken);
		}

		public void STARTERRORTRAPPINGANDCLEARANYERROR(int errorToken)
		{
			// Note: Whenever error trapping is explicitly enabled or disabled, any error is cleared. If two methods are called within an OERN..
			//   ON ERROR RESUME Next
			//   F1()
			//   F2()
			// .. and F1() raises an error, that error's information will be maintained while F2 is called (if it is called without an error being
			// raised) unless F2 or any code it calls contains On Error Resume Next or On Error Goto - if this is the case then the error from F1
			// is lost forever. This is why _trappedErrorIfAny is set to null here and in STOPERRORTRAPPINGANDCLEARANYERROR.
			if (!_activeErrorTokens.ContainsKey(errorToken))
				throw new Exception("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");
			_activeErrorTokens[errorToken] = ErrorTokenState.OnErrorResumeNext;
			_trappedErrorIfAny = null;
		}

		public void STOPERRORTRAPPINGANDCLEARANYERROR(int errorToken)
		{
			if (!_activeErrorTokens.ContainsKey(errorToken))
				throw new Exception("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");
			_activeErrorTokens[errorToken] = ErrorTokenState.OnErrorGoto0;
			_trappedErrorIfAny = null;
		}

		public void HANDLEERROR(int errorToken, Action action)
		{
			if (!_activeErrorTokens.ContainsKey(errorToken))
				throw new Exception("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");

			try
			{
				action();
			}
			catch (Exception e)
			{
				// Translated programs shouldn't provide any actions that register or unregister error tokens, but since we've just gone off and
				// attempted to do some unknown work, it's best to check
				if (!_activeErrorTokens.ContainsKey(errorToken))
					throw new Exception("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");

				if (_activeErrorTokens[errorToken] == ErrorTokenState.OnErrorResumeNext)
					SETERROR(e);
				else
				{
					RELEASEERRORTRAPPINGTOKEN(errorToken);
					throw;
				}
			}
		}

		/// <summary>
		/// This layers error-handling on top of the IAccessValuesUsingVBScriptRules.IF method, if error-handling is enabled for the specified
		/// token then evaluation of the value will be attempted - if an error occurs then it will be recorded and the condition will be treated
		/// as true, since this is VBScript's behaviour. It will throw an exception for a null valueEvaluator or an invalid errorToken.
		/// </summary>
		public bool IF(Func<object> valueEvaluator, int errorToken)
		{
			if (valueEvaluator == null)
				throw new ArgumentNullException("valueEvaluator");

			// VBScript's behaviour is quite mad here; if error-trapping is enabled when an IF condition must be evaluated, and if that evaluation results in
			// and error being raised, then act as if the condition was met. So we default to true and then try to perform the evalaluation with HANDLEERROR.
			// If an error is thrown and error-trapping is enabled, then true will be returned. If an error is throw an error-trapping is NOT enabled, then
			// that error will be allowed to propagate up. If there is no error raised then the result of the IF evaluation is returned.
			// - Note: In http://blogs.msdn.com/b/ericlippert/archive/2004/08/19/error-handling-in-vbscript-part-one.aspx, Eric Lippert does sort of
			//   describe this in passing (see the note that reads "If Blah raises an error then it resumes on the Print "Hello" in either case")
			var result = true;
			HANDLEERROR(
				errorToken,
				() => { result = _valueRetriever.IF(valueEvaluator()); }
			);
			return result;
		}

		/// <summary>
		/// This is used by implementation of CINT, CSNG, CDBL and the like - it handles special cases of types such as Empty or booleans (and with error cases
		/// such as blanks string or VBScript Null) to try to extract a number. This number will be passed through the specified converter to ensure that it is
		/// translated into the desired type. If there are no applicable special cases then the value will be passed through the VAL function and then through
		/// the processor (if this fails then a TypeMismatchException will be raised).
		/// </summary>
		private T GetAsNumber<T>(object value, string optionalExceptionMessageForInvalidContent, Func<object, T> converter) where T : struct
		{
			if (converter == null)
				throw new ArgumentNullException("nonSpecialCaseProcessor");

			value = _valueRetriever.VAL(value, optionalExceptionMessageForInvalidContent);
			value = _valueRetriever.NUM(value);
			if (value is DateTime)
				value = DateToDouble((DateTime)value);
			if (value is T)
				return (T)value;
			try
			{
				return converter(value);
			}
			catch (OverflowException e)
			{
				throw new VBScriptOverflowException(Convert.ToDouble(value), e);
			}
			catch (Exception e)
			{
				throw new TypeMismatchException(optionalExceptionMessageForInvalidContent, e);
			}
		}

		/// <summary>
		/// Given a set of values, this will return nullable ints for each of the values - null if the value was a VBScript null (ie. DBNull.Value) and an int
		/// otherwise. Along with this set, it will return a lambda which will transform an int into the largest common bitwise-applicable value type that was
		/// encountered across the values (if all of the values were booleans, then this will transform an int into a boolean, if all of the values were booleans
		/// or bytes then it will transform an int into a byte - the range of types are, in ascending order: boolean, byte, Int16, Int32). If it is is not possible
		/// to translate any of the values into an int (or if there were no values specified) then an exception will be raised. This is because the VBScript "logical"
		/// operators actually perform bitwise operations, limiting the size of those numbers to int aka Int32 aka VBScript "Long" (so any number that won't fit into
		/// the range of an Int32 will result in an overflow). After VBScript performs the operation, it will return a value that relates to the inputs - so if two
		/// booleans were operated on then a boolean will be returned, if an "Integer" (Int16) or a "Long" (Int32) were operated on then an Int32 will be returned
		/// (this is what the lambda is for).
		/// </summary>
		private Tuple<IEnumerable<int?>, Func<int, object>> GetForBitwiseOperations(string exceptionMessageForInvalidContent, params object[] values)
		{
			if (values == null)
				throw new ArgumentNullException("values");
			if (values.Length == 0)
				throw new ArgumentException("At least one value must be specified");
			if (string.IsNullOrWhiteSpace(exceptionMessageForInvalidContent))
				throw new ArgumentException("Null/blank exceptionMessageForInvalidContent specified");

			// 1. Ensure that all values are of acceptable types (note that Empty will be parsed as a number, becoming an Int32 since it has no explicit type)
			//    and DBNull.Value will remain as DBNull.Value
			values = values.Select(v => _valueRetriever.VAL(v, exceptionMessageForInvalidContent)).ToArray();

			// 2. Determine the return type based upon all of the values types and generate a lambda that will transform an Int32 into this type
			//    - It seems that VBScript does not do anything as simple as choosing the smallest data type (boolean, byte, Int16, Int32).. while in MOST cases it does
			//      that (eg. boolean and Int16 => Int16, boolean and Int32 => Int32, Int16 and Int32 => Int32, boolean and Int32 => Int32) if there is a boolean and a
			//      byte then it jumps to Int16. In fairness, I imagine this is because byte is an unsigned type and so can not represent -1, which is what boolean True
			//      is represented by as a number - so Int16 is the smallest type that can contain all boolean AND byte values.
			Func<int, object> returnTypeConverter;
			if (values.All(v => v is bool))
				returnTypeConverter = finalValue => (finalValue != 0);
			else if (values.All(v => v is byte))
				returnTypeConverter = finalValue => Convert.ToByte(finalValue & byte.MaxValue);
			else if (values.All(v => (v is bool) || (v is byte) || (v is Int16)))
			{
				// This is the only complicated type conversion, really. To translate from an Int32 into an Int16, we want to take the last 8 (of 16) bits. For
				// example, if this is part of a NOT operation that is given an Int16 value of one, then that will be translated into a long (see below) and then
				// manipulated by the caller - in this case, changing from binary "0000000000000001" to "1111111111111110" (-2 in decimal). If we mask out the last
				// 8 bits then we get "11111110" and need only cast that to an Int16. This is why we can't use Convert.ToInt16 - since the long value we have manipulated
				// could easily cause an overflow (as it would in the example here).
				returnTypeConverter = finalValue => (Int16)(finalValue & 0xffff);
			}
			else
				returnTypeConverter = finalValue => finalValue;

			// 3. Return the values as null (where VBScript Null values were found) or as Int32 values - using the convention of a C# null for a VBScript null
			//    (despite the fact that they're not the same elsewhere VBScript Null = DBNull.Value in C#, VBScript Empty = null in C#) allows us to take
			//    advantage of the Nullable<int> type, rather than having to return IEnumerable<object> where everything is either Int32 or DBNull.Value
			return Tuple.Create(
				values.Select(v => (v == DBNull.Value) ? (int?)null : CLNG(v, exceptionMessageForInvalidContent)),
				returnTypeConverter
			);
		}

		private bool IsDotNetNumericType(object l)
		{
			if (l == null)
				return false;
			return
				IsDotNetIntegerType(l) ||
				(l is decimal) || (l is double) || (l is float);
		}

		private bool IsDotNetIntegerType(object l)
		{
			if (l == null)
				return false;
			if (l.GetType().IsEnum)
				return true;
			return (l is byte) || (l is char) || (l is int) || (l is long) || (l is sbyte) || (l is short) || (l is uint) || (l is ulong) || (l is ushort);
		}

		/// <summary>
		/// The comparison (o == VBScriptConstants.Nothing) will return false even if o is VBScriptConstants.Nothing due to the implementation details of
		/// DispatchWrapper. This method delivers a reliable way to test for it.
		/// </summary>
		private bool IsVBScriptNothing(object o)
		{
			return ((o is DispatchWrapper) && ((DispatchWrapper)o).WrappedObject == null);
		}

		private double DateToDouble(DateTime value)
		{
			// When VBScript describes a date as a number, it applies somewhat counter-intuitive handling to the date component; 400.2 and -400.2 both
			// represent that same time (but on different days). Which means that -400.2 comes AFTER -400.0 chronologically, where as -400.2 comes
			// BEFORE -400 on the number scale. This behaviour needs to be reflected when we translate back from a DateTime to a double, the time
			// needs to be applied differently to a positive value than a negative. -400.2 is equivalent to "25/11/1898 04:48:00". If this date was
			// naively translated back into a number (by taking the total number of days, both whole and fractional, between the value and VBScript's
			// "zero date") then it would become -399.8. Instead the date component must be used to calculate -400 and then this SUBTRACTED from the
			// value for negatives, rather than added, so -400 becomes -400.2 (a subtraction of 0.2 from -400).
			double valueDouble;
			if (value < VBScriptConstants.ZeroDate)
				valueDouble = value.Date.Subtract(VBScriptConstants.ZeroDate).Subtract(value.TimeOfDay).TotalDays;
			else
				valueDouble = value.Date.Subtract(VBScriptConstants.ZeroDate).Add(value.TimeOfDay).TotalDays;
			return valueDouble;
		}

		/// <summary>
		/// VBScript has comparisons that will return true, false or Null (meaning DBNull.Value) which is a return type that is difficult to represent
		/// without resorting to "object" (which could be anything) or an enum (which wouldn't be the end of the world). I think the best approach,
		/// though, is to return a nullable bool from methods internally and then translate this for VBScript (so null becomes DBNull.Value).
		/// The same approach works for other non-nullable types.
		/// </summary>
		private static object ToVBScriptNullable<T>(T? value) where T : struct
		{
			if (value == null)
				return DBNull.Value;
			return value.Value;
		}

		// Feed all of these straight through to the _valueRetriever we have
		public IBuildCallArgumentProviders ARGS
		{
			get { return _valueRetriever.ARGS; }
		}
		public object CALL(object context, object target, IEnumerable<string> members, IProvideCallArguments argumentProvider)
		{
			return _valueRetriever.CALL(context, target, members, argumentProvider);
		}
		public void SET(object valueToSetTo, object context, object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider)
		{
			_valueRetriever.SET(valueToSetTo, context, target, optionalMemberAccessor, argumentProvider);
		}
		public bool IsVBScriptValueType(object o)
		{
			return _valueRetriever.IsVBScriptValueType(o);
		}
		public bool TryVAL(object o, out bool parameterLessDefaultMemberWasAvailable, out object asValueType)
		{
			return _valueRetriever.TryVAL(o, out parameterLessDefaultMemberWasAvailable, out asValueType);
		}
		public object VAL(object o, string optionalExceptionMessageForInvalidContent = null)
		{
			return _valueRetriever.VAL(o, optionalExceptionMessageForInvalidContent);
		}
		public object OBJ(object o, string optionalExceptionMessageForInvalidContent = null)
		{
			return _valueRetriever.OBJ(o, optionalExceptionMessageForInvalidContent);
		}
		public bool BOOL(object o, string optionalExceptionMessageForInvalidContent = null)
		{
			return _valueRetriever.BOOL(o, optionalExceptionMessageForInvalidContent);
		}
		public object NUM(object o, params object[] numericValuesTheTypeMustBeAbleToContain)
		{
			return _valueRetriever.NUM(o, numericValuesTheTypeMustBeAbleToContain);
		}
		public object NullableNUM(object o)
		{
			return _valueRetriever.NullableNUM(o);
		}
		public object NullableDATE(object o)
		{
			return _valueRetriever.NullableDATE(o);
		}
		public DateTime DATE(object o, string optionalExceptionMessageForInvalidContent = null)
		{
			return _valueRetriever.DATE(o);
		}
		public object NullableSTR(object o)
		{
			return _valueRetriever.NullableSTR(o);
		}
		public string STR(object o, string optionalExceptionMessageForInvalidContent = null)
		{
			return _valueRetriever.STR(o);
		}
		public bool IF(object o)
		{
			return _valueRetriever.IF(o);
		}
		public IEnumerable ENUMERABLE(object o)
		{
			return _valueRetriever.ENUMERABLE(o);
		}

		/// <summary>
		/// Where date literals were present in the source code, in a format that does not specify a date, they must be translated into dates at
		/// runtime. They must all be expanded to have whatever year it was when the request started - if the request happens to take sufficient
		/// time that the year ticks over during processing, all date literals (without explicit years) must be associated with the year when
		/// the request started. Note that if a new request starts with a different year, then date literals without years within that request
		/// must be associatedw ith the new year (this is consistent with how the VBScript interpreter would re-process the script each time).
		/// </summary>
		public DateParser DateLiteralParser { get; private set; }
	}
}
