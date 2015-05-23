using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class STRING
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object numberOfTimesToRepeat, object character, string expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object numberOfTimesToRepeat, object character)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object numberOfTimesToRepeat, object character)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object numberOfTimesToRepeat, object character)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object numberOfTimesToRepeat, object character)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character);
                });
            }

            [Theory, MemberData("InvalidProcedureCallOrArgumentData")]
            public void InvalidProcedureCallOrArgumentCases(string description, object numberOfTimesToRepeat, object character)
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character);
                });
            }

            [Theory, MemberData("OutOfStringSpaceData")]
            public void OutOfStringSpaceCases(string description, object numberOfTimesToRepeat, object character)
            {
                Assert.Throws<OutOfStringSpaceException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().STRING(numberOfTimesToRepeat, character);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty character (with numberOfTimesToRepeat 1)", 1, null, "\0" };
                    yield return new object[] { "Empty numberOfTimesToRepeat (with character \"a\")", null, "a", "" };
                    yield return new object[] { "Single character \"a\"", 1, "a", "a" };
                    yield return new object[] { "5x character \"a\"", 5, "a", "aaaaa" };
                    yield return new object[] { "5x character 65", 5, 65, new string((char)65, 5) };
                    yield return new object[] { "5x character 6500", 5, 6500, new string((char)100, 5) }; // 6500 % 256 = 100
                    yield return new object[] { "5x character -6500", 5, -6500, new string((char)156, 5) }; // -6500 + (26 * 256) = 156
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null character (with numberOfTimesToRepeat 1)", 1, DBNull.Value };
                    yield return new object[] { "Null numberOfTimesToRepeat (with character \"a\")", DBNull.Value, "a" };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Empty Array character (with numberOfTimesToRepeat 1)", 1, new object[0] };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing character (with numberOfTimesToRepeat 1)", 1, VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing numberOfTimesToRepeat (with character \"a\")", VBScriptConstants.Nothing, "a" };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "VBScript Long (.net Int32) MaxValue + 1 numberOfTimesToRepeat (with character \"a\")", (Int64)int.MaxValue + 1, "a" };
                    yield return new object[] { "VBScript Int (.net Int16) MaxValue + 1 character (with numberOfTimesToRepeat 1)", 1, (Int32)Int16.MaxValue+ 1 };
                }
            }

            public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
            {
                get
                {
                    yield return new object[] { "Blank string character (with numberOfTimesToRepeat 1)", 1, "" };
                }
            }

            public static IEnumerable<object[]> OutOfStringSpaceData
            {
                get
                {
                    yield return new object[] { "More characters than VBScript can handle (int.MaxValue / 2)", int.MaxValue / 2, "*" };
                }
            }
        }
    }
}
