using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Attributes;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ISNUMERIC
        {
            [Theory, MemberData("TrueData")]
            public void TrueCases(string description, object value)
            {
                Assert.True(DefaultRuntimeSupportClassFactory.Get().ISNUMERIC(value));
            }

            [Theory, MemberData("FalseData")]
            public void FalseCases(string description, object value)
            {
                Assert.False(DefaultRuntimeSupportClassFactory.Get().ISNUMERIC(value));
            }

            public static IEnumerable<object[]> TrueData
            {
                get
                {
                    yield return new object[] { "Empty", null };
                    yield return new object[] { "Zero", 0 };
                    yield return new object[] { "\"1\"", "1" };
                    yield return new object[] { "\"-1\"", "-1" };
                    yield return new object[] { "\"1.1\"", "1.1" };
                    yield return new object[] { "\" 1.1 \"", " 1.1 " };
                    yield return new object[] { "\" .1 \"", " .1 " };
                    yield return new object[] { "\" -.1 \"", " -.1 " };
                    yield return new object[] { "\" - .1 \"", " - .1 " };
                    yield return new object[] { "\" - 0.1 \"", " - .1 " };
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype() };
                }
            }

            public static IEnumerable<object[]> FalseData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Blank string", "" };
                    yield return new object[] { "Whitespace", " " };
                    yield return new object[] { "Whitespace around decimal point", "1. 1" };
                    yield return new object[] { "Multiple decimal points", "1..1" };
                    yield return new object[] { "Version number", "1.1.1" };
                    yield return new object[] { "Unintialised array", new object[0] };
                    yield return new object[] { "Date", new DateTime(2015, 5, 9, 13, 55, 0) };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value } };
                }
            }
        }
    }
}
