using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class MOD
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object l, object r, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().MOD(l, r));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object l, object r)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MOD(l, r);
                });
            }

            [Theory, MemberData("DivisionByZeroData")]
            public void DivisionByZeroCases(string description, object l, object r)
            {
                Assert.Throws<VBScriptDivisionByZeroException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MOD(l, r);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object l, object r)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MOD(l, r);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object l, object r)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MOD(l, r);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // Deal with the Null cases (the Empty cases are further down and in the DivisionByZeroData)
                    yield return new object[] { "Null Mod Null", DBNull.Value, DBNull.Value, DBNull.Value };
                    yield return new object[] { "1 Mod Null", 1, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null Mod 1", DBNull.Value, 1, DBNull.Value };
                    yield return new object[] { "Null Mod 0", DBNull.Value, 0, DBNull.Value }; // The Null takes precedence over the division-by-zero
                    yield return new object[] { "Null Mod \"a\"", DBNull.Value, "a", DBNull.Value }; // The Null takes precedence over the invalid string value

                    // Test the rounding cases
                    yield return new object[] { "8 Mod 7.5 (7.5 should round up to 8)", 8, 7.5, 0 };
                    yield return new object[] { "8 Mod 8.5 (8.5 should round down to 8)", 8, 8.5, 0 };
                    yield return new object[] { "7.5 Mod 10 (7.5 should round up to 8)", 7.5, 10, 8 };
                    yield return new object[] { "8.5 Mod 10 (8.5 should round down to 8)", 8.5, 10, 8 };

                    // Test the negative value cases
                    yield return new object[] { "-1 Mod 3", -1, 3, -1 };
                    yield return new object[] { "-2 Mod 3", -2, 3, -2 };
                    yield return new object[] { "-3 Mod 3", -3, 3, 0 };
                    yield return new object[] { "-4 Mod 3", -4, 3, -1 };
                    yield return new object[] { "-5 Mod 3", -5, 3, -2 };
                    yield return new object[] { "1 Mod -3", 1, -3, 1 };
                    yield return new object[] { "2 Mod -3", 2, -3, 2 };
                    yield return new object[] { "3 Mod -3", 3, -3, 0 };
                    yield return new object[] { "4 Mod -3", 4, -3, 1 };
                    yield return new object[] { "5 Mod -3", 5, -3, 2 };

                    // A Byte is only returned if both input values are Bytes. Note that the largest returnable type is Long (Int32) - Double, Currency, Date all have
                    // to be forced into a Long (and nothing larger than a Long is ever returned; in fact, onlyNull, Byte, Integer and Loner are the only possible
                    // return types)
                    yield return new object[] { "Empty Mod CByte(3)", null, (byte)3, (Int16)0 };
                    yield return new object[] { "CByte(1) Mod CByte(3)", (byte)1, (byte)3, (byte)1 };
                    yield return new object[] { "CBool(0) Mod CByte(3)", false, (byte)3, (Int16)0 };
                    yield return new object[] { "CInt(1) Mod CByte(3)", (Int16)1, (byte)3, (Int16)1 };
                    yield return new object[] { "CInt(256) Mod CByte(3)", (Int16)256, (byte)3, (Int16)1 };
                    yield return new object[] { "CLng(1) Mod CByte(3)", 1, (byte)3, 1 };
                    yield return new object[] { "CDbl(1) Mod CByte(3)", 1d, (byte)3, 1 };
                    yield return new object[] { "CDate(1) Mod CByte(3)", VBScriptConstants.ZeroDate.AddDays(1), (byte)3, 1 };
                    yield return new object[] { "CCur(1) Mod CByte(3)", 1m, (byte)3, 1 };

                    // If both values are Boolean, Byte or Integer (Int16), then the return value will be an Integer. This is also the case if the first value is Empty
                    // (if the second value is Empty then it's a division-by-zero case)
                    yield return new object[] { "CBool(1) Mod CInt(3)", true, (Int16)3, (Int16)(-1) }; // True gets treated as -1
                    yield return new object[] { "CByte(1) Mod CInt(3)", (byte)1, (Int16)3, (Int16)1 };
                    yield return new object[] { "CInt(1) Mod CInt(3)", (Int16)1, (Int16)3, (Int16)1 };
                    yield return new object[] { "CInt(1) Mod CBool(1)", (Int16)1, true, (Int16)0 };
                    yield return new object[] { "CInt(1) Mod CByte(3)", (Int16)1, (byte)3, (Int16)1 };
                    yield return new object[] { "Empty Mod CInt(3)", null, (Int16)3, (Int16)0 };

                    // Everything else results in a Long (Int32) being returned
                    yield return new object[] { "CLng(1) Mod CLng(3)", 1, 3, 1 };
                    yield return new object[] { "CDbl(1) Mod CLng(3)", 1d, 3, 1 };
                    yield return new object[] { "CDate(1) Mod CLng(3)", VBScriptConstants.ZeroDate.AddDays(1), 3, 1 };
                    yield return new object[] { "CCur(1) Mod CLng(3)", 1m, 3, 1 };
                    yield return new object[] { "CLng(1) Mod CDbl(3)", 1, 3d, 1 };
                    yield return new object[] { "CLng(1) Mod CDate(3)", 1, VBScriptConstants.ZeroDate.AddDays(3), 1 };
                    yield return new object[] { "CLng(1) Mod CCur(3)", 1, 3m, 1 };
                    yield return new object[] { "CDbl(1) Mod CDbl(3)", 1d, 3d, 1 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "\"\" Mod 3", "a", 3 };
                    yield return new object[] { "\"a\" Mod 3", "a", 3 };
                    yield return new object[] { "\"2015-03-02\" Mod 3", "2015-03-02", 3 };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "(Int32.MaxValue + 1) Mod 0", 2147483648, 0 }; // The overflow takes precedence over the division-by-zero
                }
            }

            public static IEnumerable<object[]> DivisionByZeroData
            {
                get
                {
                    yield return new object[] { "Empty Mod Empty", null, null };
                    yield return new object[] { "CBool(1) Mod Empty", true, null };
                    yield return new object[] { "CByte(1) Mod Empty", (byte)1, null };
                    yield return new object[] { "CInt(1) Mod Empty", (Int16)1, null };
                    yield return new object[] { "CLng(1) Mod Empty", 1, null };
                    yield return new object[] { "CDbl(1) Mod Empty", 1d, null };
                    yield return new object[] { "CDate(1) Mod Empty", VBScriptConstants.ZeroDate.AddDays(1), null };
                    yield return new object[] { "CCur(1) Mod Empty", 1m, null };

                    yield return new object[] { "CInt(1) Mod CBool(0)", (Int16)1, false };
                    yield return new object[] { "CInt(1) Mod CByte(0)", (Int16)1, (byte)0 };
                    yield return new object[] { "CInt(1) Mod CInt(0)", (Int16)1, (Int16)0 };
                    yield return new object[] { "CInt(1) Mod CLng(0)", (Int16)1, 0 };
                    yield return new object[] { "CInt(1) Mod CDbl(0)", (Int16)1, 0d};
                    yield return new object[] { "CInt(1) Mod CDate(0)", (Int16)1, VBScriptConstants.ZeroDate };
                    yield return new object[] { "CInt(1) Mod CCur(0)", (Int16)1, 0m };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    // Null values can mask a division-by-zero but not object-variable-not-set
                    yield return new object[] { "Null Mod Nothing", DBNull.Value, VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing Mod Null", VBScriptConstants.Nothing, DBNull.Value };
                }
            }
        }
    }
}
