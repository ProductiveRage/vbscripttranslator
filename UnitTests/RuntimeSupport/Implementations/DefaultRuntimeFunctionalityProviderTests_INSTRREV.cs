using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class INSTRREV // Note that INSTRREV has a different method signature to INSTR (the startIndex argument is in a different place)
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object valueToSearch, object valueToSearchFor, object startIndex, object compareMode, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().INSTRREV(valueToSearch, valueToSearchFor, startIndex, compareMode));
            }

            [Theory, MemberData("SuccessDataWithNoStartIndex")]
            public void SuccessCasesWithoutStartIndexValues(string description, object valueToSearch, object valueToSearchFor, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().INSTRREV(valueToSearch, valueToSearchFor));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object valueToSearch, object valueToSearchFor, object startIndex, object compareMode)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().INSTRREV(valueToSearch, valueToSearchFor, startIndex, compareMode);
                });
            }

            [Theory, MemberData("InvalidProcedureCallOrArgumentData")]
            public void InvalidProcedureCallOrArgumentCases(string description, object valueToSearch, object valueToSearchFor, object startIndex, object compareMode)
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().INSTRREV(valueToSearch, valueToSearchFor, startIndex, compareMode);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object valueToSearch, object valueToSearchFor, object startIndex, object compareMode)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().INSTRREV(valueToSearch, valueToSearchFor, startIndex, compareMode);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // These are edge cases that return values (and so are "successes" in that they do not result in exception being raised)
                    yield return new object[] { "Empty valueToSearch returns no match", null, "test", 1, 0, 0 };
                    yield return new object[] { "Empty valueToSearchFor returns immediate match (for non-Empty valueToSearch)", "test", null, 1, 0, 1 };
                    yield return new object[] { "Empty valueToSearchFor returns no match for Empty valueToSearch", null, null, 1, 0, 0 };
                    yield return new object[] { "Null valueToSearch returns Null", DBNull.Value, "test", 1, 0, DBNull.Value };
                    yield return new object[] { "Null valueToSearchFor returns Null", "test", DBNull.Value, 1, 0, DBNull.Value };

                    // These are more traditional success cases
                    yield return new object[] { "Not found returns zero", "test", "z", 1, 0, 0 };
                    yield return new object[] { "Match on last character returns length of valueToSearch if valueToSearchFor is one character long", "test", "t", 4, 0, 4 };
                    yield return new object[] { "CompareMode 1 means match case-insensitive", "TEST", "t", 4, 1, 4 };
                    yield return new object[] { "StartIndex n-1 skips the last character", "test", "t", 3, 1, 1 };

                    // These are more edge cases that nevertheless return real values
                    yield return new object[] { "Standard rounding rules are applied to startIndex (0.9 rounds to 1)", "Test", "t", 0.9, 1, 1 };
                    yield return new object[] { "Standard rounding rules are applied to compareMode (1.4 rounds to 1)", "Test", "t", 1, 1.4, 1 };
                    yield return new object[] { "If valueToSearchFor is longer than valueToSearch then zero is returned", "Test", "aaaaaaaaa", 1, 0, 0 };
                    yield return new object[] { "If startIndex is larger than valueToSearch's length then zero is returned", "Test", "t", 5, 0, 0 };

                    // Some more traditional success cases
                    yield return new object[] { "'t' in 'tttt' working back from character 5 returns 0", "tttt", "t", 5, 0, 0 }; // startIndex > valueToSearch.length => 0 returned
                    yield return new object[] { "'t' in 'tttt' working back from character 4 returns 1", "tttt", "t", 4, 0, 4 };
                    yield return new object[] { "'t' in 'tttt' working back from character 3 returns 1", "tttt", "t", 3, 0, 3 };
                    yield return new object[] { "'t' in 'tttt' working back from character 2 returns 1", "tttt", "t", 2, 0, 2 };
                    yield return new object[] { "'t' in 'tttt' working back from character 1 returns 1", "tttt", "t", 1, 0, 1 };
                    yield return new object[] { "'t' in 'etttt' working back from character 6 returns 0", "etttt", "t", 6, 0, 0 }; // startIndex > valueToSearch.length => 0 returned
                    yield return new object[] { "'t' in 'etttt' working back from character 5 returns 1", "etttt", "t", 5, 0, 5 };
                    yield return new object[] { "'t' in 'etttt' working back from character 4 returns 1", "etttt", "t", 4, 0, 4 };
                    yield return new object[] { "'t' in 'etttt' working back from character 3 returns 1", "etttt", "t", 3, 0, 3 };
                    yield return new object[] { "'t' in 'etttt' working back from character 2 returns 2", "etttt", "t", 2, 0, 2 };
                    yield return new object[] { "'t' in 'etttt' working back from character 1 has no match", "etttt", "t", 1, 0, 0 };
                }
            }

            public static IEnumerable<object[]> SuccessDataWithNoStartIndex
            {
                get
                {
                    // If no startIndex is specified then it should default to the length of valueToSearch, unless that can not be evaluated as a non-blank string (in which
                    // case it will default to a minimum of one - or result in a type exception)
                    yield return new object[] { "'test' has default startIndex of 4", "test", "t", 4 };
                    yield return new object[] { "Default startIndex for '' is irrelevant (but does not error)", null, "t", 0 };
                    yield return new object[] { "Default startIndex for Empty is irrelevant (but does not error)", null, "t", 0 };
                    yield return new object[] { "Default startIndex for Null is irrelevant (but does not error)", DBNull.Value, "t", DBNull.Value };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null startIndex results in invalid-use-of-null", "search in", "search for", DBNull.Value, 0 };
                    yield return new object[] { "Null compareMode results in invalid-use-of-null", "search in", "search for", 1, DBNull.Value };
                }
            }

            public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
            {
                get
                {
                    yield return new object[] { "Zero startIndex results in invalid-procedure-call-or-argument", "search in", "search for", 0, 0 };
                    yield return new object[] { "Negative startIndex results in invalid-procedure-call-or-argument", "search in", "search for", -1, 0 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Non-numeric-or-empty-or-null startIndex results in type-mismatch", "search in", "search for", "x", 0 };
                    yield return new object[] { "Non-numeric-or-empty-or-null compareMode results in type-mismatch", "search in", "search for", 1, "x" };
                }
            }
        }
    }
}
