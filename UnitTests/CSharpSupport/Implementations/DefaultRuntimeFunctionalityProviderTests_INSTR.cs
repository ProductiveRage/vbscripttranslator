using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class INSTR
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object startIndex, object valueToSearch, object valueToSearchFor, object compareMode, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().INSTR(startIndex, valueToSearch, valueToSearchFor, compareMode));
            }
            
            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object startIndex, object valueToSearch, object valueToSearchFor, object compareMode)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().INSTR(startIndex, valueToSearch, valueToSearchFor, compareMode);
                });
            }

            [Theory, MemberData("InvalidProcedureCallOrArgumentData")]
            public void InvalidProcedureCallOrArgumentCases(string description, object startIndex, object valueToSearch, object valueToSearchFor, object compareMode)
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().INSTR(startIndex, valueToSearch, valueToSearchFor, compareMode);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object startIndex, object valueToSearch, object valueToSearchFor, object compareMode)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().INSTR(startIndex, valueToSearch, valueToSearchFor, compareMode);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // These are edge cases that return values (and so are "successes" in that they do not result in exception being raised)
                    yield return new object[] { "Empty valueToSearch returns no match", 1, null, "test", 0, 0 };
                    yield return new object[] { "Empty valueToSearchFor returns immediate match (for non-Empty valueToSearch)", 1, "test", null, 0, 1 };
                    yield return new object[] { "Empty valueToSearchFor returns no match for Empty valueToSearch", 1, null, null, 0, 0 };
                    yield return new object[] { "Null valueToSearch returns Null", 1, DBNull.Value, "test", 0, DBNull.Value };
                    yield return new object[] { "Null valueToSearchFor returns Null", 1, "test", DBNull.Value, 0, DBNull.Value };

                    // These are more traditional success cases
                    yield return new object[] { "Not found returns zero", 1, "test", "z", 0, 0 };
                    yield return new object[] { "Match on first characters returns 1 (1-based index)", 1, "test", "t", 0, 1 };
                    yield return new object[] { "CompareMode 1 means match case-insensitive", 1, "Test", "t", 1, 1 };
                    yield return new object[] { "StartIndex 2 skips the first character", 2, "test", "t", 1, 4 };
                    yield return new object[] { "StartIndex 4 will return 4 when searching for 't' within 'test'", 4, "test", "t", 0, 4 }; // Off-by-one check

                    // These are more edge cases that nevertheless return real values
                    yield return new object[] { "Standard rounding rules are applied to startIndex (0.9 rounds to 1)", 0.9, "Test", "t", 1, 1 };
                    yield return new object[] { "Standard rounding rules are applied to compareMode (1.4 rounds to 1)", 1, "Test", "t", 1.4, 1 };
                    yield return new object[] { "If valueToSearchFor is longer than valueToSearch then zero is returned", 1, "Test", "aaaaaaaaa", 0, 0 };
                    yield return new object[] { "If startIndex is larger than valueToSearch's length then zero is returned", 5, "Test", "t", 0, 0 };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null startIndex results in invalid-use-of-null", DBNull.Value, "search in", "search for", 0 };
                    yield return new object[] { "Null compareMode results in invalid-use-of-null", 1, "search in", "search for", DBNull.Value };
                }
            }

            public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
            {
                get
                {
                    yield return new object[] { "Zero startIndex results in invalid-procedure-call-or-argument", 0, "search in", "search for", 0 };
                    yield return new object[] { "Negative startIndex results in invalid-procedure-call-or-argument", -1, "search in", "search for", 0 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Non-numeric-or-empty-or-null startIndex results in type-mismatch", "x", "search in", "search for", 0 };
                    yield return new object[] { "Non-numeric-or-empty-or-null compareMode results in type-mismatch", 1, "search in", "search for", "x" };
                }
            }
        }
    }
}
