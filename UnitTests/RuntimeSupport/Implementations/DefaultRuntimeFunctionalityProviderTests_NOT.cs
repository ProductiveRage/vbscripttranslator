using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class NOT
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().NOT(value));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NOT(value);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object value)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NOT(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().NOT(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, -1 };
                    yield return new object[] { "Null", DBNull.Value, DBNull.Value };
                    yield return new object[] { "True", true, false };
                    yield return new object[] { "False", false, true };
                    yield return new object[] { "Zero (Int16)", (Int16)0, (Int16)(-1) };
                    yield return new object[] { "Minus one (Int16)", (Int16)(-1), (Int16)0 };
                    yield return new object[] { "Int16.MaxValue", Int16.MaxValue, Int16.MinValue };
                    yield return new object[] { "One (Int16)", (Int16)1, (Int16)(-2) };
                    yield return new object[] { "Zero (Byte)", (byte)0, (byte)255 };
                    yield return new object[] { "One (Byte)", (byte)1, (byte)254 };
                    yield return new object[] { "255 (Byte)", (byte)255, (byte)0 };
                    yield return new object[] { "Int32.MinValue", Int32.MinValue, Int32.MaxValue }; // Smallest value before overflow
                    yield return new object[] { "Int32.MaxValue", Int32.MaxValue, Int32.MinValue }; // Largest value before overflow

                    yield return new object[] { "0.1", 0.1, -1 }; // 0.1 will be rounded down to 0 and treated as a VBScript Long since it is not explicitly a Boolean, Byte or Integer
                    yield return new object[] { "0.5", 0.5, -1 }; // 0.5 will be rounded down to 0 and so be the same 0.1
                    yield return new object[] { "1.1", 1.1, -2 }; // 1.1 will be rounded down to 1
                    yield return new object[] { "1.5", 1.5, -3 }; // 1.5 will be rounded up to 2
                    yield return new object[] { "2 (Int16)", (Int16)2, (Int16)(-3) }; // 1.5 will be rounded up to 2

                    yield return new object[] { "String \"1.1\"", "1.1", -2 }; // The string "1.1" will be parsed into the number 1.1
                    yield return new object[] { "String \"1\"", "1", -2 }; // The string "1" will be parsed into the number 1 (and treated as a VBScript Long)

                    yield return new object[] { "Date 2015-05-28 16:04:58", new DateTime(2015, 5, 28, 16, 4, 58), -42154 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Blank string", "" };
                    yield return new object[] { "1D array", new object[0] };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "Int32.MinValue - 1", (Int64)Int32.MinValue - 1 };
                    yield return new object[] { "Int32.MaxValue + 1", (Int64)Int32.MaxValue + 1 };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                }
            }
        }
    }
}
