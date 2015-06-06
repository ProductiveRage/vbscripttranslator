using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ADD
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object l, object r, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().ADD(l, r));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object l, object r)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().ADD(l, r);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object l, object r)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().ADD(l, r);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object l, object r)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().ADD(l, r);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // Empty is treated as an Integer (Int16) zero
                    yield return new object[] { "Empty + Empty", null, null, (Int16)0 };

                    // Addition with Null will always return Null (unless the other value is Nothing)
                    yield return new object[] { "Empty + Null", null, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null + Empty", DBNull.Value, null, DBNull.Value };
                    yield return new object[] { "Null + Null", DBNull.Value, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null + 0", DBNull.Value, (Int16)0, DBNull.Value };
                    yield return new object[] { "0 + Null", (Int16)0, DBNull.Value, DBNull.Value };

                    // Booleans are only returned if both inputs are booleans
                    yield return new object[] { "CBool(0) + CBool(0)", false, false, false };
                    yield return new object[] { "CBool(0) + CBool(1)", false, true, true };
                    yield return new object[] { "CBool(1) + CBool(0)", true, false, true };
                    yield return new object[] { "CBool(1) + CBool(1)", true, true, (Int16)(-2) }; // This is an overflow, CBool(1) => (Int16)(-1)
                    yield return new object[] { "CBool(0) + Empty", false, null, (Int16)0 }; // Empty is treated as (Int16)0 and so this the inputs are not both booleans and so a boolean is not returned
                    yield return new object[] { "CBool(1) + Empty", true, null, (Int16)(-1) };
                    yield return new object[] { "CBool(1) + Null", true, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CBool(1) + CByte(3)", true, (byte)3, (Int16)2 }; // Byte can not contain all values of Boolean (0 and -1) so the next type up for them both is Integer (Int16)
                    yield return new object[] { "CBool(0) + CByte(3)", false, (byte)3, (Int16)3 }; // Similar to above data but to prove Bool + Byte => Integer is based on types and not values
                    yield return new object[] { "CBool(1) + CInt(3)", true, (Int16)3, (Int16)2 }; // CBool(1) becomes (Int16)(-1) for this operation so the result is (Int16)2
                    yield return new object[] { "CBool(1) + CLng(3)", true, 3, 2 };
                    yield return new object[] { "CBool(1) + CDbl(3)", true, 3d, 2d };
                    yield return new object[] { "CBool(1) + CCur(3)", true, 3m, 2m };
                    yield return new object[] { "CBool(1) + CDate(3)", true, VBScriptConstants.ZeroDate.AddDays(3), VBScriptConstants.ZeroDate.AddDays(2) };

                    // Bytes are only returned if both inputs are bytes or the non-byte input is Empty (and if the result would not overflow byte's range)
                    yield return new object[] { "CByte(0) + CByte(0)", (byte)0, (byte)0, (byte)0 };
                    yield return new object[] { "CByte(0) + CByte(1)", (byte)0, (byte)1, (byte)1 };
                    yield return new object[] { "CByte(1) + CByte(0)", (byte)1, (byte)0, (byte)1 };
                    yield return new object[] { "CByte(0) + Empty", (byte)0, null, (byte)0 }; // I'm surprised that this isn't an Integer return, since Empty is (Int16)0 when added to a Boolean.. but the result here IS supposed to be of type Byte
                    yield return new object[] { "CByte(1) + Empty", (byte)1, null, (byte)1 };
                    yield return new object[] { "Empty + CByte(1)", null, (byte)1, (byte)1 };
                    yield return new object[] { "CByte(1) + Null", (byte)1, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CByte(255) + CByte(1)", (byte)255, (byte)1, (Int16)256 }; // Overflows into an Int16
                    yield return new object[] { "CByte(1) + CInt(3)", (byte)1, (Int16)3, (Int16)4 };
                    yield return new object[] { "CByte(1) + CLng(3)", (byte)1, 3, 4 };
                    yield return new object[] { "CByte(1) + CDbl(3)", (byte)1, 3d, 4d };
                    yield return new object[] { "CByte(1) + CCur(3)", (byte)1, 3m, 4m };
                    yield return new object[] { "CByte(1) + CDate(3)", (byte)1, VBScriptConstants.ZeroDate.AddDays(3), VBScriptConstants.ZeroDate.AddDays(4) };

                    // Currency is more resistant to changing type that other numbers, it will not change type to avoid an overflow. The return value will be Currency unless
                    // the other value is Null or if it is a Date (note that if it IS a date then it WILL overflow, if required, into a Double)
                    yield return new object[] { "CCur(1) + CCur(1)", 1m, 1m, 2m };
                    yield return new object[] { "CCur(1) + Empty", 1m, null, 1m };
                    yield return new object[] { "CCur(1) + Null", 1m, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CCur(1) + CBool(1)", 1m, true, 0m }; // CBool(1) becomes -1, so the total is zero (type remains Currency)
                    yield return new object[] { "CCur(1) + CByte(1)", 1m, (byte)1, 2m };
                    yield return new object[] { "CCur(1) + CInt(1)", 1m, (int)1, 2m };
                    yield return new object[] { "CCur(1) + CLng(1)", 1m, 1, 2m };
                    yield return new object[] { "CCur(1) + CDbl(1)", 1m, 1d, 2m };
                    yield return new object[] { "CCur(1) + CDate(1)", 1m, VBScriptConstants.ZeroDate.AddDays(1), VBScriptConstants.ZeroDate.AddDays(2) }; // Type changes to date
                    yield return new object[] { "CCur(value-bigger-than-date-can-describe) + CDate(1)", 10000000m, VBScriptConstants.ZeroDate.AddDays(1), 10000001d }; // Type would change to date but overflows to double

                    // Operations on Date values also want to return Date results, but (unlike Currency) they overflow into Double if required
                    yield return new object[] { "CDate(1) + CDate(1)", VBScriptConstants.ZeroDate.AddDays(1), VBScriptConstants.ZeroDate.AddDays(1), VBScriptConstants.ZeroDate.AddDays(2) };
                    yield return new object[] { "CDate(1) + CBool(1)", VBScriptConstants.ZeroDate.AddDays(1), true, VBScriptConstants.ZeroDate }; // CBool(1) is treated as -1
                    yield return new object[] { "CDate(1) + CInt(1)", VBScriptConstants.ZeroDate.AddDays(1), (Int16)1, VBScriptConstants.ZeroDate.AddDays(2) };
                    yield return new object[] { "CDate(1) + CLng(1)", VBScriptConstants.ZeroDate.AddDays(1), 1, VBScriptConstants.ZeroDate.AddDays(2) };
                    yield return new object[] { "CDate(1) + CDbl(1)", VBScriptConstants.ZeroDate.AddDays(1), 1d, VBScriptConstants.ZeroDate.AddDays(2) };
                    yield return new object[] { "CDate(1) + Empty", VBScriptConstants.ZeroDate.AddDays(1), null, VBScriptConstants.ZeroDate.AddDays(1) };
                    yield return new object[] { "CDate(1) + Null", VBScriptConstants.ZeroDate.AddDays(1), DBNull.Value, DBNull.Value };
                    yield return new object[] { "CDate(max-date-value) + 1", VBScriptConstants.LatestPossibleDate, 1, VBScriptConstants.LatestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays + 1 };
                    yield return new object[] { "CDate(min-date-value) + (-1)", VBScriptConstants.EarliestPossibleDate, -1, VBScriptConstants.EarliestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays - 1 };

                    // Integer is straight-forward; it will move up a type if the other value is of a large size and will move up to a Long if it would otherwise overflow
                    yield return new object[] { "CInt(1) + CInt(1)", (Int16)1, (Int16)1, (Int16)2 };
                    yield return new object[] { "CInt(32767) + CInt(1)", (Int16)32767, (Int16)1, 32768 }; // Overflows into a VBScript Long (C# Int32)
                    yield return new object[] { "CInt(-32768) + CInt(-1)", (Int16)(-32768), (Int16)(-1), -32769 }; // Overflows into a VBScript Long (C# Int32)
                    yield return new object[] { "CInt(1) + Empty", (Int16)1, null, (Int16)1 }; // Empty is treated as Integer (Int16) zero
                    yield return new object[] { "CInt(1) + Null", (Int16)1, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CInt(1) + CLng(1)", (Int16)1, 1, 2 };
                    yield return new object[] { "CInt(1) + CDbl(1)", (Int16)1, 1d, 2d };
                    yield return new object[] { "CInt(1) + CDate(1)", (Int16)1, VBScriptConstants.ZeroDate.AddDays(1), VBScriptConstants.ZeroDate.AddDays(2) };
                    yield return new object[] { "CInt(1) + CCur(1)", (Int16)1, 1m, 2m };

                    // Long is very similar to Integer
                    yield return new object[] { "CLng(1) + CLng(1)", 1, 1, 2 };
                    yield return new object[] { "CLng(2147483647) + CLng(1)", 2147483647, 1, 2147483648d }; // Overflows into a Double
                    yield return new object[] { "CLng(-2147483648) + CLng(-1)", -2147483648, -1, -2147483649d }; // Overflows into a Double
                    yield return new object[] { "CLng(1) + Empty", 1, null, 1 }; // Empty is treated as Integer (Int16) zero
                    yield return new object[] { "CLng(1) + Null", 1, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CLng(1) + CDbl(1)", 1, 1d, 2d };
                    yield return new object[] { "CLng(1) + CDate(1)", 1, VBScriptConstants.ZeroDate.AddDays(1), VBScriptConstants.ZeroDate.AddDays(2) };
                    yield return new object[] { "CLng(1) + CCur(1)", 1, 1m, 2m };

                    // Most of the Double cases have been handled above..
                    yield return new object[] { "CDbl(1) + CDbl(1)", 1d, 1d, 2d };
                    yield return new object[] { "CDbl(1) + Empty", 1d, null, 1d };
                    yield return new object[] { "CDbl(1) + Null", 1d, DBNull.Value, DBNull.Value };

                    // Strings have a few rules:
                    // - If both values are strings then the addition operates as a concatenation
                    // - If a string is "added" to Empty then Empty is interpreted as a blank string
                    // - If a string is "added" to Null then Null is returned (as with other values)
                    // - If a string is "added" to a numeric (or boolean) value, the string must be numeric and will then be treated as a Double
                    yield return new object[] { "\"1\" + \"2\"", "1", "2", "12" };
                    yield return new object[] { "\"1\" + Empty", "1", null, "1" };
                    yield return new object[] { "Empty + \"1\"", null, "1", "1" };
                    yield return new object[] { "\"1\" + Null", "1", DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null + \"1\"", DBNull.Value, "1", DBNull.Value };
                    yield return new object[] { "CInt(1) + \"1\"", (Int16)1, "1", 2d };
                    yield return new object[] { "CBool(1) + \"3\"", true, "3", 2d }; // CBool(1) / true => (Int16)(-1), "3" => 3d, so the result is 2d

                    // The standard if-object-then-try-to-access-default-parameterless-function-or-property logic applies
                    yield return new object[] { "Object-with-default-property-with-value-Integer-1 + CInt(2)", new exampledefaultpropertytype { result = (Int16)1 }, (Int16)2, (Int16)3 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    // Blank and non-numeric values are invalid (the number parsing logic is shared with the NUM function, so the tests there cover the variety of acceptable and
                    // unacceptable values more thoroughly)
                    yield return new object[] { "CInt(1) + \"\"", (Int16)1, "" };
                    yield return new object[] { "CInt(1) + \"a\"", (Int16)1, "a" };

                    // String representations of boolean values are not considered valid for addition
                    yield return new object[] { "CInt(1) + \"True\"", (Int16)1, "True" };
                    yield return new object[] { "CInt(1) + \"true\"", (Int16)1, "true" };
                    
                    // String representations of dates values are not considered valid for addition
                    yield return new object[] { "CInt(1) + \"2015-03-02\"", (Int16)1, "2015-03-02" };

                    // No wrangling to arrays is supported (not even "if it's one-dimensional and has only a single element then use that")
                    yield return new object[] { "CInt(1) + Array()", (Int16)1, new object[0] };
                    yield return new object[] { "CInt(1) + Array(1)", (Int16)1, new object[] { 1 } };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "VBScript.MaxCurrency + 1", VBScriptConstants.MaxCurrencyValue, 1 };
                    yield return new object[] { "VBScript.MinCurrency + (-1)", VBScriptConstants.MinCurrencyValue, -1 };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing + Empty", VBScriptConstants.Nothing, null };
                    yield return new object[] { "Empty + Nothing", null, VBScriptConstants.Nothing };

                    yield return new object[] { "Nothing + Null", VBScriptConstants.Nothing, DBNull.Value };
                    yield return new object[] { "Null + Nothing", VBScriptConstants.Nothing, DBNull.Value };

                    yield return new object[] { "Nothing + 1", VBScriptConstants.Nothing, 1 };

                    yield return new object[] { "Object-without-default-member + 1", new Object(), 1 };
                }
            }
        }
    }
}
