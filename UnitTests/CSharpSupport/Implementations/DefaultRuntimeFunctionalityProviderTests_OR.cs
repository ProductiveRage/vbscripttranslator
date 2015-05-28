using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class OR
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object l, object r, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().OR(l, r));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object l, object r)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().OR(l, r);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object l, object r)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().OR(l, r);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object l, object r)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().OR(l, r);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty OR Empty", null, null, 0 }; // Result is type VBScript "Long" (ie. Int32) since no type is explicitly stated where Empty is used

                    // If only one value is Null, then the other value will be returned (but if both are Null then Null will be returned)
                    yield return new object[] { "Null OR Null", DBNull.Value, DBNull.Value, DBNull.Value };
                    yield return new object[] { "1 OR Null", 1, DBNull.Value, 1 };
                    yield return new object[] { "Null OR 1", DBNull.Value, 1, 1 };

                    yield return new object[] { "True OR True", true, true, true };
                    yield return new object[] { "True OR False", true, false, true };
                    yield return new object[] { "False OR True", false, true, true };
                    yield return new object[] { "False OR False", false, false, false };

                    yield return new object[] { "CByte(1) OR CByte(2)", (byte)1, (byte)2, (byte)3 };
                    yield return new object[] { "CByte(1) OR CByte(3)", (byte)1, (byte)3, (byte)3 };
                    yield return new object[] { "CByte(1) OR CInt(2)", (byte)1, (Int16)2, (Int16)3 };
                    yield return new object[] { "CByte(1) OR CInt(3)", (byte)1, (Int16)3, (Int16)3 };

                    // Note: -1 | 1 = -1 since -1 is 11111111111111111111111111111111 (when an Int32), since it's the negative number furthest away from the most negative value (which is Int32.MinValue
                    // or 10000000000000000000000000000000) while 1 is all bits off other than the last one. When these are OR'd together, all of the bits stay on, which represents -1.
                    yield return new object[] { "CDbl(1) OR CDbl(-1)", 1m, -1m, -1 }; // Double is treated as Int32 since decimals are not supported in bitwise operations

                    yield return new object[] { "CByte(1) OR True", (byte)1, true, (Int16)(-1) }; // This is the same calculation as 1 | -1
                    yield return new object[] { "CByte(1) OR False", (byte)1, false, (Int16)1 }; // The smallest type that contain contain byte values (0-255) and boolean values (0 or -1) is Int16

                    yield return new object[] { "CInt(1) OR True", (Int16)1, true, (Int16)(-1) }; // This is 1 | -1 again
                    yield return new object[] { "CInt(1) OR False", (Int16)1, false, (Int16)1 };
                    yield return new object[] { "CInt(1) OR CByte(0)", (Int16)1, (byte)0, (Int16)1 };
                    yield return new object[] { "CInt(1) OR CByte(1)", (Int16)1, (byte)1, (Int16)1 };
                    yield return new object[] { "CInt(1) OR CByte(25)", (Int16)1, (byte)25, (Int16)25 };

                    yield return new object[] { "CLng(1) OR True", 1, true, -1 }; // This is 1 | -1 again again
                    yield return new object[] { "CLng(1) OR False", 1, false, 1 };

                    // Largest value before overflow
                    yield return new object[] { "Int32.MaxValue OR 2", int.MaxValue, 2, int.MaxValue };
                    yield return new object[] { "Int32.MaxValue OR Null", int.MaxValue, DBNull.Value, int.MaxValue }; // Unlike AND, if only one value is Null then the non-null value is returned

                    // Smallest value before negative overflow
                    // - Note: int.MinValue is "10000000000000000000000000000000" (zero places above the most negative value) and -2 is "11111111111111111111111111111110" (a LOT of places above the
                    //   most negative value), so when they're OR'd together, it results in "11111111111111111111111111111110", which is still -2
                    yield return new object[] { "Int32.MinValue OR -2", int.MinValue, -2, -2 };
                    yield return new object[] { "Int32.MinValue OR Null", int.MinValue, DBNull.Value, int.MinValue };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Blank string OR Empty", "", null };
                    yield return new object[] { "Blank string OR Null", "", DBNull.Value };
                    yield return new object[] { "Blank string OR Blank string", "", "" };
                    yield return new object[] { "Blank string OR 1", "", 1 };
                    
                    yield return new object[] { "1D array OR Empty", new object[0], null };
                    yield return new object[] { "1D array OR Null", new object[0], DBNull.Value };
                    yield return new object[] { "1D array OR 1D array", new object[0], new object[0] };
                    yield return new object[] { "1D array OR 1", new object[0], 1 };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "(Int32.MinValue - 1) OR 0", (Int64)Int32.MinValue - 1, 0 };
                    yield return new object[] { "(Int32.MaxValue + 1) OR 0", (Int64)Int32.MaxValue + 1, 0 };
                    
                    // If either value is VBScript Null then VBScript Null will be returned, so long as every value can be evaluated as a number within the allowable range
                    yield return new object[] { "(Int32.MinValue - 1) OR Null", (Int64)Int32.MinValue - 1, DBNull.Value };
                    yield return new object[] { "(Int32.MaxValue + 1) OR Null", (Int64)Int32.MaxValue + 1, DBNull.Value };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing OR Empty", VBScriptConstants.Nothing, null };
                    yield return new object[] { "Nothing OR Null", VBScriptConstants.Nothing, DBNull.Value };
                    yield return new object[] { "Nothing OR Nothing", VBScriptConstants.Nothing, VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing OR 1", VBScriptConstants.Nothing, 1 };
                }
            }
        }
    }
}
