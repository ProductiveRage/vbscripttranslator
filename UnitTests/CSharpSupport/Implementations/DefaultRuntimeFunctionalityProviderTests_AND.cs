using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class AND
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object l, object r, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().AND(l, r));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object l, object r)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().AND(l, r);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object l, object r)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().AND(l, r);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object l, object r)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().AND(l, r);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty AND Empty", null, null, 0 }; // Result is type VBScript "Long" (ie. Int32) since no type is explicitly stated where Empty is used

                    // If one or both of the values are Null, then Null will be returned
                    yield return new object[] { "Null AND Null", DBNull.Value, DBNull.Value, DBNull.Value };
                    yield return new object[] { "1 OR Null", 1, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Null OR 1", DBNull.Value, 1, DBNull.Value };

                    yield return new object[] { "True AND True", true, true, true };
                    yield return new object[] { "True AND False", true, false, false };
                    yield return new object[] { "False AND True", false, true, false };
                    yield return new object[] { "False AND False", false, false, false };

                    yield return new object[] { "CByte(1) AND CByte(2)", (byte)1, (byte)2, (byte)0 };
                    yield return new object[] { "CByte(1) AND CByte(3)", (byte)1, (byte)3, (byte)1 };
                    yield return new object[] { "CByte(1) AND CInt(2)", (byte)1, (Int16)2, (Int16)0 };
                    yield return new object[] { "CByte(1) AND CInt(3)", (byte)1, (Int16)3, (Int16)1 };

                    yield return new object[] { "CDbl(1) AND CDbl(-1)", 1m, -1m, 1 }; // Double is treated as Int32 since decimals are not supported in bitwise operations

                    yield return new object[] { "CByte(1) AND True", (byte)1, true, (Int16)1 };
                    yield return new object[] { "CByte(1) AND False", (byte)1, false, (Int16)0 }; // The smallest type that contain contain byte values (0-255) and boolean values (0 or -1) is Int16

                    yield return new object[] { "CInt(1) AND True", (Int16)1, true, (Int16)1 };
                    yield return new object[] { "CInt(1) AND False", (Int16)1, false, (Int16)0 };
                    yield return new object[] { "CInt(1) AND CByte(0)", (Int16)1, (byte)0, (Int16)0 };
                    yield return new object[] { "CInt(1) AND CByte(1)", (Int16)1, (byte)1, (Int16)1 };
                    yield return new object[] { "CInt(1) AND CByte(25)", (Int16)1, (byte)25, (Int16)1 };

                    yield return new object[] { "CLng(1) AND True", 1, true, 1 };
                    yield return new object[] { "CLng(1) AND False", 1, false, 0 };

                    // Largest value before overflow
                    yield return new object[] { "Int32.MaxValue AND 2", int.MaxValue, 2, 2 };
                    yield return new object[] { "Int32.MaxValue AND Null", int.MaxValue, DBNull.Value, DBNull.Value };

                    // Smallest value before negative overflow
                    // - Note: The first result caught me out, why does MinValue AND -2 equal MinValue?! But it makes sense since int.MinValue is represented by "10000000000000000000000000000000"
                    //   (meaning zero places above the most negative value, int.MinValue + 1 is "10000000000000000000000000000001"; one place above the most negative value) and since -2 is
                    //   "11111111111111111111111111111110" (a LOT of places about the most negative value). When "10000000000000000000000000000000" and "11111111111111111111111111111110"
                    //   are bitwise AND'd, only the first bit stays is one and the rest are zero - which is exactly the same as int.MinValue's binary value!
                    yield return new object[] { "Int32.MinValue AND -2", int.MinValue, -2, int.MinValue };
                    yield return new object[] { "Int32.MinValue AND Null", int.MinValue, DBNull.Value, DBNull.Value };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Blank string AND Empty", "", null };
                    yield return new object[] { "Blank string AND Null", "", DBNull.Value };
                    yield return new object[] { "Blank string AND Blank string", "", "" };
                    yield return new object[] { "Blank string AND 1", "", 1 };
                    
                    yield return new object[] { "1D array AND Empty", new object[0], null };
                    yield return new object[] { "1D array AND Null", new object[0], DBNull.Value };
                    yield return new object[] { "1D array AND 1D array", new object[0], new object[0] };
                    yield return new object[] { "1D array AND 1", new object[0], 1 };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "(Int32.MinValue - 1) AND 0", (Int64)Int32.MinValue - 1, 0 };
                    yield return new object[] { "(Int32.MaxValue + 1) AND 0", (Int64)Int32.MaxValue + 1, 0 };
                    
                    // If either value is VBScript Null then VBScript Null will be returned, so long as every value can be evaluated as a number within the allowable range
                    yield return new object[] { "(Int32.MinValue - 1) AND Null", (Int64)Int32.MinValue - 1, DBNull.Value };
                    yield return new object[] { "(Int32.MaxValue + 1) AND Null", (Int64)Int32.MaxValue + 1, DBNull.Value };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing AND Empty", VBScriptConstants.Nothing, null };
                    yield return new object[] { "Nothing AND Null", VBScriptConstants.Nothing, DBNull.Value };
                    yield return new object[] { "Nothing AND Nothing", VBScriptConstants.Nothing, VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing AND 1", VBScriptConstants.Nothing, 1 };
                }
            }
        }
    }
}
