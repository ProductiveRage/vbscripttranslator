using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class MULT
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object l, object r, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().MULT(l, r));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object l, object r)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MULT(l, r);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object l, object r)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MULT(l, r);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object l, object r)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MULT(l, r);
                });
            }

            [Theory, MemberData("ObjectDoesNotSupportPropertyOrMemberData")]
            public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object l, object r)
            {
                Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MULT(l, r);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // Empty is treated as an Integer (Int16) zero
                    yield return new object[] { "Empty * Empty", null, null, (Int16)0 };

                    // Multiplication with Null will always return Null (unless the other value is Nothing)
                    yield return new object[] { "Empty * Null", null, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null * Empty", DBNull.Value, null, DBNull.Value };
                    yield return new object[] { "Null * Null", DBNull.Value, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null * 0", DBNull.Value, (Int16)0, DBNull.Value };
                    yield return new object[] { "0 * Null", (Int16)0, DBNull.Value, DBNull.Value };

                    // The multiplication of two Booleans always results in Integers. Same for one Boolean with an Empty or with a Byte. Same with an Integer, unless it would overflow (which
                    // can only happen with true, which is converted to -1, and the most negative Integer value, since its positive value is outside of the Integer range). The same happens
                    // with Long (the result is long with one exception). Double and Currency always pull the result up to their own type when multiplied by a Boolean. Dates multiplied
                    // by Booleans always result in a Double.
                    yield return new object[] { "CBool(0) * CBool(0)", false, false, (Int16)0 };
                    yield return new object[] { "CBool(0) * CBool(1)", false, true, (Int16)0 };
                    yield return new object[] { "CBool(1) * CBool(0)", true, false, (Int16)0 };
                    yield return new object[] { "CBool(1) * CBool(1)", true, true, (Int16)1 }; // CBool(1) will be treated as (Int16)(-1), so the result is -1 x -1 = 1
                    yield return new object[] { "CBool(0) * CBool(0)", false, false, (Int16)0 };
                    yield return new object[] { "CBool(0) * Empty", false, null, (Int16)0 };
                    yield return new object[] { "CBool(1) * Empty", true, null, (Int16)0 };
                    yield return new object[] { "CBool(0) * Null", false, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CBool(0) * CByte(0)", false, (byte)0, (Int16)0 };
                    yield return new object[] { "CBool(0) * CInt(0)", false, (Int16)0, (Int16)0 };
                    yield return new object[] { "CBool(1) * CInt(-32768)", true, (Int16)(-32768), 32768 }; // The only Boolean / Integer multiplication that results in a Long instead of an Integer
                    yield return new object[] { "CBool(0) * CLng(0)", false, (Int32)0, (Int32)0 };
                    yield return new object[] { "CBool(1) * CLng(-2147483648)", true, -2147483648, 2147483648d }; // The only Boolean / Long multiplication that results in a Double instead of a Long
                    yield return new object[] { "CBool(0) * CDbl(0)", false, 0d, 0d };
                    yield return new object[] { "CBool(0) * CCur(0)", false, 0m, 0m };
                    yield return new object[] { "CBool(0) * CDate(0)", false, VBScriptConstants.ZeroDate, 0d };

                    // Multiplying two Bytes will result in a Byte, unless it would overflow, in which case it will return an Integer. Multiplying with Empty will return a Byte, though multiplying
                    // with an Integer will result in an Integer - as will Long, Double and Currency). Multiplying with a Date will result in a Double.
                    yield return new object[] { "CByte(0) * CByte(0)", (byte)0, (byte)0, (byte)0 };
                    yield return new object[] { "CByte(2) * CByte(2)", (byte)2, (byte)2, (byte)4 };
                    yield return new object[] { "CByte(16) * CByte(16)", (byte)16, (byte)16, (Int16)256 }; // Overflow into an Integer
                    yield return new object[] { "CByte(16) * Empty", (byte)16, null, (byte)0 };
                    yield return new object[] { "CByte(16) * Null", (byte)16, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CByte(16) * CInt(1)", (byte)16, (Int16)1, (Int16)16 };
                    yield return new object[] { "CByte(16) * CLng(1)", (byte)16, 1, 16 };
                    yield return new object[] { "CByte(16) * CDbl(1)", (byte)16, 1d, 16d };
                    yield return new object[] { "CByte(16) * CCur(1)", (byte)16, 1m, 16m };
                    yield return new object[] { "CByte(16) * CDate(1)", (byte)16, VBScriptConstants.ZeroDate.AddDays(1), 16d };

                    // No surprises with Integer - move up type to prevent overflow or if the other value is of a larger type
                    yield return new object[] { "CInt(0) * CInt(0)", (Int16)0, (Int16)0, (Int16)0 };
                    yield return new object[] { "CInt(2) * CInt(2)", (Int16)2, (Int16)2, (Int16)4 };
                    yield return new object[] { "CInt(256) * CInt(256)", (Int16)256, (Int16)256, 65536 }; // Overflow into an Long
                    yield return new object[] { "CInt(16) * Empty", (Int16)16, null, (Int16)0 };
                    yield return new object[] { "CInt(16) * Null", (Int16)16, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CInt(16) * CLng(1)", (Int16)16, 1, 16 };
                    yield return new object[] { "CInt(16) * CDbl(1)", (Int16)16, 1d, 16d };
                    yield return new object[] { "CInt(16) * CCur(1)", (Int16)16, 1m, 16m };
                    yield return new object[] { "CInt(16) * CDate(1)", (Int16)16, VBScriptConstants.ZeroDate.AddDays(1), 16d };

                    // Same with Long..
                    yield return new object[] { "CLng(0) * CLng(0)", 0, 0, 0 };
                    yield return new object[] { "CLng(2) * CLng(2)", 2, 2, 4 };
                    yield return new object[] { "CLng(50000) * CLng(50000)", 50000, 50000, 2500000000d }; // Overflow into a Double
                    yield return new object[] { "CLng(16) * Empty", 16, null, 0 };
                    yield return new object[] { "CLng(16) * Null", 16, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CLng(16) * CDbl(1)", 16, 1d, 16d };
                    yield return new object[] { "CLng(16) * CCur(1)", 16, 1m, 16m };
                    yield return new object[] { "CLng(16) * CDate(1)", 16, VBScriptConstants.ZeroDate.AddDays(1), 16d };

                    // A Currency multiplied by a Double or a Date will result in a Double, but otherwise it tries to stick to the Currency type - even preferring to raise an error on overflow
                    // rather than moving up to the Double type
                    yield return new object[] { "CCur(0) * CCur(0)", 0m, 0m, 0m };
                    yield return new object[] { "CCur(2) * CCur(2)", 2m, 2m, 4m };
                    yield return new object[] { "CCur(16) * Empty", 16m, null, 0m };
                    yield return new object[] { "CCur(16) * Null", 16m, DBNull.Value, DBNull.Value };
                    yield return new object[] { "CCur(16) * CDbl(1)", 16m, 1d, 16d };
                    yield return new object[] { "CCur(16) * CDate(1)", 16m, VBScriptConstants.ZeroDate.AddDays(1), 16d };

                    // Multiplying Dates always results in a Double
                    yield return new object[] { "CDate(0) * CDate(0)", VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate, 0d };
                    yield return new object[] { "CDate(16) * Empty", VBScriptConstants.ZeroDate.AddDays(16), null, 0d };
                    yield return new object[] { "CDate(16) * Null", VBScriptConstants.ZeroDate.AddDays(16), DBNull.Value, DBNull.Value };

                    // Double * Double => Double
                    yield return new object[] { "CDbl(0) * CDbl(0)", 0d, 0d, 0d };
                    yield return new object[] { "CDbl(0) * Empty", 0d, null, 0d };
                    yield return new object[] { "CDbl(0) * Null", 0d, DBNull.Value, DBNull.Value };

                    // Strings are interpreted as Doubles if they're numeric (error case if they're not - including where they COULD be parsed as boolean or dates; that's not allowed)
                    yield return new object[] { "\"2\" * CByte(1)", "2", (byte)1, 2d };
                    yield return new object[] { "\"2\" * CInt(1)", "2", (Int16)1, 2d };
                    yield return new object[] { "\"2\" * CLng(1)", "2", 1, 2d };
                    yield return new object[] { "\"2\" * CCur(1)", "2", 1m, 2d };
                    yield return new object[] { "\"2\" * CDbl(1)", "2", 1d, 2d };
                    yield return new object[] { "\"-2\" * CDbl(1)", "-2", 1d, -2d };
                    yield return new object[] { "\"0.2\" * CDbl(1)", "0.2", 1d, 0.2d };
                    yield return new object[] { "\"2\" * Empty", "2", null, 0d };
                    yield return new object[] { "\"2\" * Null", "2", DBNull.Value, DBNull.Value };

                    // The standard if-object-then-try-to-access-default-parameterless-function-or-property logic applies
                    yield return new object[] { "Object-with-default-property-with-value-Integer-1 * CInt(2)", new exampledefaultpropertytype { result = (Int16)1 }, (Int16)2, (Int16)2 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    // Blank and non-numeric values are invalid (the number parsing logic is shared with the NUM function, so the tests there cover the variety of acceptable and
                    // unacceptable values more thoroughly)
                    yield return new object[] { "CInt(1) * \"\"", (Int16)1, "" };
                    yield return new object[] { "CInt(1) * \"a\"", (Int16)1, "a" };

                    // String representations of boolean values are not considered valid for multiplication
                    yield return new object[] { "CInt(1) * \"True\"", (Int16)1, "True" };
                    yield return new object[] { "CInt(1) * \"true\"", (Int16)1, "true" };

                    // String representations of dates values are not considered valid for multiplication
                    yield return new object[] { "CInt(1) * \"2015-03-02\"", (Int16)1, "2015-03-02" };

                    // No wrangling to arrays is supported (not even "if it's one-dimensional and has only a single element then use that")
                    yield return new object[] { "CInt(1) * Array()", (Int16)1, new object[0] };
                    yield return new object[] { "CInt(1) * Array(1)", (Int16)1, new object[] { 1 } };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "CCur(40000000) * CCur(40000000)", 40000000m, 40000000m }; // This could fit into a Double, but Currency won't move up types to avoid overflow
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing * Empty", VBScriptConstants.Nothing, null };
                    yield return new object[] { "Empty * Nothing", null, VBScriptConstants.Nothing };

                    yield return new object[] { "Nothing * Null", VBScriptConstants.Nothing, DBNull.Value };
                    yield return new object[] { "Null * Nothing", DBNull.Value, VBScriptConstants.Nothing };

                    yield return new object[] { "Nothing * 1", VBScriptConstants.Nothing, 1 };
                }
            }

            public static IEnumerable<object[]> ObjectDoesNotSupportPropertyOrMemberData
            {
                get
                {
                    yield return new object[] { "Object-without-default-member * 1", new Object(), 1 };
                }
            }
        }
    }
}
