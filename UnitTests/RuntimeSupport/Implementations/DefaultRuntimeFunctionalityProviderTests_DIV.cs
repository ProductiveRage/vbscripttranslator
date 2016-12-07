using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		public class DIV
		{
			[Theory, MemberData("SuccessData")]
			public void SuccessCases(string description, object l, object r, object expectedResult)
			{
				Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().DIV(l, r));
			}

			[Theory, MemberData("TypeMismatchData")]
			public void TypeMismatchCases(string description, object l, object r)
			{
				Assert.Throws<TypeMismatchException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DIV(l, r);
				});
			}

			[Theory, MemberData("OverflowData")]
			public void OverflowCases(string description, object l, object r)
			{
				Assert.Throws<VBScriptOverflowException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DIV(l, r);
				});
			}

			[Theory, MemberData("ObjectVariableNotSetData")]
			public void ObjectVariableNotSetCases(string description, object l, object r)
			{
				Assert.Throws<ObjectVariableNotSetException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DIV(l, r);
				});
			}

			[Theory, MemberData("ObjectDoesNotSupportPropertyOrMemberData")]
			public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object l, object r)
			{
				Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DIV(l, r);
				});
			}

			public static IEnumerable<object[]> SuccessData
			{
				get
				{
					// Empty is treated as zero
					yield return new object[] { "Empty / 1", null, (short)1, 0.0 };

					// Division with Null will always return Null (unless the other value is Nothing)
					yield return new object[] { "Empty / Null", null, DBNull.Value, DBNull.Value };
					yield return new object[] { "Null / Empty", DBNull.Value, null, DBNull.Value };
					yield return new object[] { "Null / Null", DBNull.Value, DBNull.Value, DBNull.Value };
					yield return new object[] { "Null / 0", DBNull.Value, (short)0, DBNull.Value };
					yield return new object[] { "0 / Null", (short)0, DBNull.Value, DBNull.Value };

					// These are all the cases that lead to a Single being returned (instead of a Double)
					// One side of the expression needs to be a Single, and the other side must be either Single, Integer (Int16), Byte, or Boolean
					// In addition, if the result of the calculation would overflow a Single, the value is promoted to a Double
					yield return new object[] { "CByte(5) / CSng(2)", (byte)5, 2.0f, 2.5f };
					yield return new object[] { "CInt(5) / CSng(2)", (short)5, 2.0f, 2.5f };
					yield return new object[] { "CSng(5) / CByte(2)", 5.0f, (byte)2, 2.5f };
					yield return new object[] { "CSng(5) / CInt(2)", 5.0f, (short)2, 2.5f };
					yield return new object[] { "CSng(5) / CSng(2)", 5.0f, 2.0f, 2.5f };

					// A calculation using Singles typically return a single, but if the result overflows a Single, it'll get promoted to Double
					yield return new object[] { "CSng(3.402823e+38) / CSng(1)", 3.402823e+38f, 1.0f, 3.402823e+38f }; // Fits inside a Single, remains a Single
					yield return new object[] { "CSng(3.402823e+38) / CSng(0.1)", 3.402823e+38f, 0.1f, 3.4028230100310822e+39 }; // Overflows a Single, promoted to Double

					// Any other cases will return a Double value
					yield return new object[] { "CByte(5) / CByte(2)", (byte)5, (byte)2, 2.5 };
					yield return new object[] { "CByte(5) / CInt(2)", (byte)5, (short)2, 2.5 };
					yield return new object[] { "CByte(5) / CLng(2)", (byte)5, 2, 2.5 };
					yield return new object[] { "CByte(5) / CDbl(2)", (byte)5, 2.0, 2.5 };
					yield return new object[] { "CByte(5) / CCur(2)", (byte)5, 2.0m, 2.5 };
					yield return new object[] { "CByte(5) / CStr(2)", (byte)5, "2", 2.5 };
					yield return new object[] { "CByte(5) / CDate(2)", (byte)5, VBScriptConstants.ZeroDate.AddDays(2), 2.5 };
					yield return new object[] { "CInt(5) / CByte(2)", (short)5, (byte)2, 2.5 };
					yield return new object[] { "CInt(5) / CInt(2)", (short)5, (short)2, 2.5 };
					yield return new object[] { "CInt(5) / CLng(2)", (short)5, 2, 2.5 };
					yield return new object[] { "CInt(5) / CDbl(2)", (short)5, 2.0, 2.5 };
					yield return new object[] { "CInt(5) / CCur(2)", (short)5, 2.0m, 2.5 };
					yield return new object[] { "CInt(5) / CStr(2)", (short)5, "2", 2.5 };
					yield return new object[] { "CInt(5) / CDate(2)", (short)5, VBScriptConstants.ZeroDate.AddDays(2), 2.5 };
					yield return new object[] { "CLng(5) / CByte(2)", 5, (byte)2, 2.5 };
					yield return new object[] { "CLng(5) / CInt(2)", 5, (short)2, 2.5 };
					yield return new object[] { "CLng(5) / CLng(2)", 5, 2, 2.5 };
					yield return new object[] { "CLng(5) / CDbl(2)", 5, 2.0, 2.5 };
					yield return new object[] { "CLng(5) / CCur(2)", 5, 2.0m, 2.5 };
					yield return new object[] { "CLng(5) / CStr(2)", 5, "2", 2.5 };
					yield return new object[] { "CLng(5) / CDate(2)", 5, VBScriptConstants.ZeroDate.AddDays(2), 2.5 };
					yield return new object[] { "CDbl(5) / CByte(2)", 5.0, (byte)2, 2.5 };
					yield return new object[] { "CDbl(5) / CInt(2)", 5.0, (short)2, 2.5 };
					yield return new object[] { "CDbl(5) / CLng(2)", 5.0, 2, 2.5 };
					yield return new object[] { "CDbl(5) / CDbl(2)", 5.0, 2.0, 2.5 };
					yield return new object[] { "CDbl(5) / CCur(2)", 5.0, 2.0m, 2.5 };
					yield return new object[] { "CDbl(5) / CStr(2)", 5.0, "2", 2.5 };
					yield return new object[] { "CDbl(5) / CDate(2)", 5.0, VBScriptConstants.ZeroDate.AddDays(2), 2.5 };
					yield return new object[] { "CCur(5) / CByte(2)", 5.0m, (byte)2, 2.5 };
					yield return new object[] { "CCur(5) / CInt(2)", 5.0m, (short)2, 2.5 };
					yield return new object[] { "CCur(5) / CLng(2)", 5.0m, 2, 2.5 };
					yield return new object[] { "CCur(5) / CDbl(2)", 5.0m, 2.0, 2.5 };
					yield return new object[] { "CCur(5) / CCur(2)", 5.0m, 2.0m, 2.5 };
					yield return new object[] { "CCur(5) / CStr(2)", 5.0m, "2", 2.5 };
					yield return new object[] { "CCur(5) / CDate(2)", 5.0m, VBScriptConstants.ZeroDate.AddDays(2), 2.5 };
					yield return new object[] { "CStr(5) / CByte(2)", "5", (byte)2, 2.5 };
					yield return new object[] { "CStr(5) / CInt(2)", "5", (short)2, 2.5 };
					yield return new object[] { "CStr(5) / CLng(2)", "5", 2, 2.5 };
					yield return new object[] { "CStr(5) / CDbl(2)", "5", 2.0, 2.5 };
					yield return new object[] { "CStr(5) / CCur(2)", "5", 2.0m, 2.5 };
					yield return new object[] { "CStr(5) / CStr(2)", "5", "2", 2.5 };
					yield return new object[] { "CStr(5) / CDate(2)", "5", VBScriptConstants.ZeroDate.AddDays(2), 2.5 };
					yield return new object[] { "CDate(5) / CByte(2)", VBScriptConstants.ZeroDate.AddDays(5), (byte)2, 2.5 };
					yield return new object[] { "CDate(5) / CInt(2)", VBScriptConstants.ZeroDate.AddDays(5), (short)2, 2.5 };
					yield return new object[] { "CDate(5) / CLng(2)", VBScriptConstants.ZeroDate.AddDays(5), 2, 2.5 };
					yield return new object[] { "CDate(5) / CDbl(2)", VBScriptConstants.ZeroDate.AddDays(5), 2.0, 2.5 };
					yield return new object[] { "CDate(5) / CCur(2)", VBScriptConstants.ZeroDate.AddDays(5), 2.0m, 2.5 };
					yield return new object[] { "CDate(5) / CStr(2)", VBScriptConstants.ZeroDate.AddDays(5), "2", 2.5 };
					yield return new object[] { "CDate(5) / CDate(2)", VBScriptConstants.ZeroDate.AddDays(5), VBScriptConstants.ZeroDate.AddDays(2), 2.5 };

					// Just checking that the maximum positive and negative Double-precision values don't error for any reason
					yield return new object[] { "CDbl(1.7976931348623157e+308) / CDbl(1)", 1.7976931348623157e+308, 1.0, 1.7976931348623157e+308 };
					yield return new object[] { "CDbl(1.7976931348623157e+308) / CDbl(-1)", 1.7976931348623157e+308, -1.0, -1.7976931348623157e+308 };
					yield return new object[] { "CDbl(-1.7976931348623157e+308) / CDbl(1)", -1.7976931348623157e+308, 1.0, -1.7976931348623157e+308 };
					yield return new object[] { "CDbl(-1.7976931348623157e+308) / CDbl(-1)", -1.7976931348623157e+308, -1.0, 1.7976931348623157e+308 };

					// The standard if-object-then-try-to-access-default-parameterless-function-or-property logic applies
					yield return new object[] { "Object-with-default-property-with-value-Double-5 / CDbl(2)", new exampledefaultpropertytype { result = 5.0 }, 2.0, 2.5 };
				}
			}

			public static IEnumerable<object[]> TypeMismatchData
			{
				get
				{
					// Blank and non-numeric values are invalid (the number parsing logic is shared with the NUM function, so the tests there cover the variety of acceptable and
					// unacceptable values more thoroughly)
					yield return new object[] { "CInt(1) / \"\"", (short)1, "" };
					yield return new object[] { "CInt(1) / \"a\"", (short)1, "a" };

					// String representations of boolean values are not considered valid for division
					yield return new object[] { "CInt(1) / \"True\"", (short)1, "True" };
					yield return new object[] { "CInt(1) / \"true\"", (short)1, "true" };

					// String representations of dates values are not considered valid for division
					yield return new object[] { "CInt(1) / \"2015-03-02\"", (short)1, "2015-03-02" };

					// No wrangling to arrays is supported (not even "if it's one-dimensional and has only a single element then use that")
					yield return new object[] { "CInt(1) / Array()", (short)1, new object[0] };
					yield return new object[] { "CInt(1) / Array(1)", (short)1, new object[] { 1 } };
				}
			}

			public static IEnumerable<object[]> OverflowData
			{
				get
				{
					// Overflow is caused when:
					//    1. The left-hand-side is zero and the right-hand-side is zero
					// OR 2. The result of the division overflows the bounds of a Double-precision floating point number
					yield return new object[] { "CDbl(0) / CDbl(0)", 0.0, 0.0 };
					yield return new object[] { "CDbl(1.7976931348623157e+308) / CDbl(0.1)", 1.7976931348623157e+308, 0.1 };
					yield return new object[] { "CDbl(-1.7976931348623157e+308) / CDbl(0.1)", -1.7976931348623157e+308, 0.1 };

					// Booleans are converted to their numeric form before division, so True is -1 and False is 0
					yield return new object[] { "False / False", false, false };
					yield return new object[] { "CDbl(0) / False", 0.0, false };
				}
			}

			public static IEnumerable<object[]> DivisionByZeroData
			{
				get
				{
					// Division-by-zero is caused when the left-hand side is non-zero and the right-hand-side is zero
					yield return new object[] { "CDbl(1) / CDbl(0)", 1.0, 0.0 };
					yield return new object[] { "CDbl(-1) / CDbl(0)", -1.0, 0.0 };

					// Booleans are converted to their numeric form before division, so True is -1 and False is 0
					yield return new object[] { "True / False", true, false };
					yield return new object[] { "CDbl(4) / False", 4.0, false };
				}
			}

			public static IEnumerable<object[]> ObjectVariableNotSetData
			{
				get
				{
					yield return new object[] { "Nothing / Empty", VBScriptConstants.Nothing, null };
					yield return new object[] { "Empty / Nothing", null, VBScriptConstants.Nothing };

					yield return new object[] { "Nothing / Null", VBScriptConstants.Nothing, DBNull.Value };
					yield return new object[] { "Null / Nothing", DBNull.Value, VBScriptConstants.Nothing };

					yield return new object[] { "Nothing / 1", VBScriptConstants.Nothing, 1 };
				}
			}

			public static IEnumerable<object[]> ObjectDoesNotSupportPropertyOrMemberData
			{
				get
				{
					yield return new object[] { "Object-without-default-member / 1", new Object(), 1 };
				}
			}
		}
	}
}
