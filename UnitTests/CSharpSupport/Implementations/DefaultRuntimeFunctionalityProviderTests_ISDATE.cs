using System;
using System.Collections.Generic;
using System.Globalization;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.UnitTests.Shared;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ISDATE : CultureOverridingTests
        {
            public ISDATE() : base(new CultureInfo("en-GB")) { }

            [Theory, MemberData("TrueData")]
            public void TrueCases(string description, object value)
            {
                Assert.True(DefaultRuntimeSupportClassFactory.Get().ISDATE(value));
            }

            [Theory, MemberData("FalseData")]
            public void FalseCases(string description, object value)
            {
                Assert.False(DefaultRuntimeSupportClassFactory.Get().ISDATE(value));
            }

            public static IEnumerable<object[]> TrueData
            {
                get
                {
                    yield return new object[] { "A DateTime", new DateTime(2015, 5, 11) };
                    yield return new object[] { "A DateTime with time component", new DateTime(2015, 5, 11, 20, 12, 44) };
                    yield return new object[] { "A 'yyyy-MM-dd' string", "2015-05-11" };
                    yield return new object[] { "A 'yyyy-M-d' string", "2015-5-11" };
                    yield return new object[] { "A 'yyyy-MM-dd HH:mm:ss' string", "2015-05-11 20:12:44" };
                    yield return new object[] { "Object with default property which is a 'yyyy-MM-dd HH:mm:ss' string", new exampledefaultpropertytype { result = "2015-05-11 20:12:44" } };

                    yield return new object[] { "String 'M d yyyy' while using en-GB culture", "1 13 2015" };
                    yield return new object[] { "String 'M yy' while using en-GB culture", "1 0" };
                }
            }

            public static IEnumerable<object[]> FalseData
            {
                get
                {
                    yield return new object[] { "Empty", null };
                    yield return new object[] { "Null", DBNull.Value };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Zero", 0 };
                    yield return new object[] { "Date-esque number", 40000 }; // Although CDate(40000) returns a valid date (2009-07-06), IsDate will return false
                    yield return new object[] { "Blank string", "" };
                    yield return new object[] { "Unintialised array", new object[0] };
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype() };
                }
            }
        }
    }
}
