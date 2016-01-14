using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class REPLACE
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode, string expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, compareMode));
            }

            [Theory, MemberData("SuccessWithMinimumArgumentsData")]
            public void SuccessWithMinimumArgumentsCases(string description, object value, object toSearchFor, object toReplaceWith, string expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, compareMode);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, compareMode);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, compareMode);
                });
            }

            [Theory, MemberData("InvalidProcedureCallOrArgumentData")]
            public void InvalidProcedureCallOrArgumentCases(string description, object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode)
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, compareMode);
                });
            }
            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements, object compareMode)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, compareMode);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\" (case-sensitive)", "testEst", "es", "{*^*}", 1, -1, 0, "t{*^*}tEst" };
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\" (case-insensitive)", "testEst", "es", "{*^*}", 1, -1, 1, "t{*^*}t{*^*}t" };
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\" (case-insensitive), zero occurrences", "testEst", "es", "{*^*}", 1, 0, 1, "testEst" };
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\" (case-insensitive), up to one occurrence", "testEst", "es", "{*^*}", 1, 1, 1, "t{*^*}tEst" };
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\" (case-insensitive), up to two occurrences", "testEst", "es", "{*^*}", 1, 2, 1, "t{*^*}t{*^*}t" };
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\" (case-insensitive), up to three occurrences", "testEst", "es", "{*^*}", 1, 3, 1, "t{*^*}t{*^*}t" };

                    yield return new object[] { "maxNumberOfReplacements int.MaxValue (not large enough to overflow)", "testEst", "es", "{*^*}", 1, int.MaxValue, 0, "t{*^*}tEst" };
                    yield return new object[] { "startIndex equal to the max allowed VBScript string length (not large enough to overflow)", "testEst", "es", "{*^*}", (int.MaxValue / 2) - 1, -1, 0, "testEst" };
                }
            }

            public static IEnumerable<object[]> SuccessWithMinimumArgumentsData
            {
                get
                {
                    yield return new object[] { "Replace empty string in \"test\" with \"Z\"", "test", "", "Z", "test" };
                    yield return new object[] { "Replace \"t\" in \"test\" with empty string", "test", "t", "", "es" };
                    yield return new object[] { "Replace \"t\" in empty string with \"Z\"", "", "t", "Z", "" };
                    yield return new object[] { "Replace \"t\" in \"test\" with \"Z\"", "test", "t", "Z", "ZesZ" };
                    yield return new object[] { "Replace \"es\" in \"testEst\" with \"{*^*}\"", "testEst", "es", "{*^*}", "t{*^*}tEst" };

                    yield return new object[] { "Replace 1 (integer) in 1 (integer) with 2 (integer)", 1, 1, 2, "2" };
                    yield return new object[] { "Replace \".\" in 1.1 (number) with \"*\"", 1.1, ".", "*", "1*1" };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null value", DBNull.Value, "es", "{*^*}", 1, -1, 0 };
                    yield return new object[] { "Null toSearchFor", "testEst", DBNull.Value, "{*^*}", 1, -1, 0 };
                    yield return new object[] { "Null toReplaceWith", "testEst", "es", DBNull.Value, 1, -1, 0 };
                    yield return new object[] { "Null startIndex", "testEst", "es", "{*^*}", DBNull.Value, -1, 0 };
                    yield return new object[] { "Null maxNumberOfReplacements", "testEst", "es", "{*^*}", 1, DBNull.Value, 0 };
                    yield return new object[] { "Null compareMode", "testEst", "es", "{*^*}", 1, -1, DBNull.Value };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Array value", new object[0], "es", "{*^*}", 1, -1, 0 };
                    yield return new object[] { "Array toSearchFor", "testEst", new object[0], "{*^*}", 1, -1, 0 };
                    yield return new object[] { "Array toReplaceWith", "testEst", "es", new object[0], 1, -1, 0 };
                    yield return new object[] { "Array startIndex", "testEst", "es", "{*^*}", new object[0], -1, 0 };
                    yield return new object[] { "Array maxNumberOfReplacements", "testEst", "es", "{*^*}", 1, new object[0], 0 };
                    yield return new object[] { "Array compareMode", "testEst", "es", "{*^*}", 1, -1, new object[0] };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing value", VBScriptConstants.Nothing, "es", "{*^*}", 1, -1, 0 };
                    yield return new object[] { "Nothing toSearchFor", "testEst", VBScriptConstants.Nothing, "{*^*}", 1, -1, 0 };
                    yield return new object[] { "Nothing toReplaceWith", "testEst", "es", VBScriptConstants.Nothing, 1, -1, 0 };
                    yield return new object[] { "Nothing startIndex", "testEst", "es", "{*^*}", VBScriptConstants.Nothing, -1, 0 };
                    yield return new object[] { "Nothing maxNumberOfReplacements", "testEst", "es", "{*^*}", 1, VBScriptConstants.Nothing, 0 };
                    yield return new object[] { "Nothing compareMode", "testEst", "es", "{*^*}", 1, -1, VBScriptConstants.Nothing };
                }
            }

            public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
            {
                get
                {
                    yield return new object[] { "compareMode 2 (must be equatable to 0 or 1)", "testEst", "es", "{*^*}", 1, -1, 2 };
                    yield return new object[] { "compareMode int.MaxValue (not large enough to overflow)", "testEst", "es", "{*^*}", 1, -1, int.MaxValue };
                    yield return new object[] { "startIndex greater than max allowed VBScript string length", "testEst", "es", "{*^*}", int.MaxValue / 2, -1, 0 };
                }
            }

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "compareMode int.MaxValue + 1 (large enough to overflow)", "testEst", "es", "{*^*}", 1, -1, (Int64)int.MaxValue + 1 };
                    yield return new object[] { "maxNumberOfReplacements int.MaxValue + 1 (large enough to overflow)", "testEst", "es", "{*^*}", 1, (Int64)int.MaxValue + 1, 0 };
                }
            }
        }
    }
}
