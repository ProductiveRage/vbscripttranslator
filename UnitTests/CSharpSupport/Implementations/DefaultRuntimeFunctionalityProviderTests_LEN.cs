using System;
using System.Collections.Generic;
using System.Globalization;
using CSharpSupport;
using VBScriptTranslator.UnitTests.Shared;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class LEN
        {
            [Fact]
            public void EmptyResultsInZero()
            {
                Assert.Equal((int)0, DefaultRuntimeSupportClassFactory.Get().LEN(null)); // This should return an int ("Long" in VBScript parlance)
            }

            [Fact]
            public void NullResultsInNull()
            {
                Assert.Equal(DBNull.Value, DefaultRuntimeSupportClassFactory.Get().LEN(DBNull.Value));
            }

            [Fact]
            public void Test()
            {
                Assert.Equal(4, DefaultRuntimeSupportClassFactory.Get().LEN("Test"));
            }

            [Fact]
            public void NumericValue()
            {
                Assert.Equal(1, DefaultRuntimeSupportClassFactory.Get().LEN(4)); // Numbers get cast as strings, so the number 4 becomes the string "4" and so has length 1
            }

            public class en_GB : CultureOverridingTests
            {
                public en_GB() : base(new CultureInfo("en-GB")) { }

                [Theory, MemberData("SuccessData")]
                public void SuccessCases(string description, object value, int expectedResult)
                {
                    Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().LEN(value));
                }

                public static IEnumerable<object[]> SuccessData
                {
                    get
                    {
                        yield return new object[] { "Date with zero time", new DateTime(2015, 5, 28), "28/05/2015".Length };
                        yield return new object[] { "Date with non-zero time", new DateTime(2015, 5, 28, 18, 54, 36), "28/05/2015 18:54:36".Length };
                        yield return new object[] { "Zero date with non-zero time", VBScriptConstants.ZeroDate.Add(new TimeSpan(18, 54, 36)), "18:54:36".Length };
                        yield return new object[] { "Zero date with zero time", VBScriptConstants.ZeroDate, "00:00:00".Length };
                    }
                }
            }
        }
    }
}
