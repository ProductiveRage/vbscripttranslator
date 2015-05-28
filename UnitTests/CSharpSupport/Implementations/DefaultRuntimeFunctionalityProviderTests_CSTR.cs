using System;
using System.Collections.Generic;
using System.Globalization;
using CSharpSupport;
using CSharpSupport.Exceptions;
using VBScriptTranslator.UnitTests.Shared;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CSTR
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, string expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().CSTR(value));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CSTR(value);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CSTR(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CSTR(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // Note: CSTR handling of dates varies by culture, so there are tests classes specifically around this further down in this file
                    yield return new object[] { "Empty", null, "" };
                    yield return new object[] { "Blank string", "", "" };
                    yield return new object[] { "Populated string", "abc", "abc" };
                    yield return new object[] { "Integer 1", 1, "1" };
                    yield return new object[] { "Floating point 1.23", 1.23, "1.23" };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value } };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "An empty array", new object[0] };
                    yield return new object[] { "Object with default property which is an empty array", new exampledefaultpropertytype { result = new object[0] } };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
                }
            }

            public class en_GB : CultureOverridingTests
            {
                public en_GB() : base(new CultureInfo("en-GB")) { }

                [Theory, MemberData("SuccessData")]
                public void SuccessCases(string description, object value, string expectedResult)
                {
                    Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().CSTR(value));
                }

                public static IEnumerable<object[]> SuccessData
                {
                    get
                    {
                        yield return new object[] { "Date with zero time", new DateTime(2015, 5, 28), "28/05/2015" };
                        yield return new object[] { "Date with non-zero time", new DateTime(2015, 5, 28, 18, 54, 36), "28/05/2015 18:54:36" };
                        yield return new object[] { "Zero date with non-zero time", VBScriptConstants.ZeroDate.Add(new TimeSpan(18, 54, 36)), "18:54:36" };
                        yield return new object[] { "Zero date with zero time", VBScriptConstants.ZeroDate, "00:00:00" };
                    }
                }
            }

            public class en_US : CultureOverridingTests
            {
                public en_US() : base(new CultureInfo("en-US")) { }

                [Theory, MemberData("SuccessData")]
                public void SuccessCases(string description, object value, string expectedResult)
                {
                    Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().CSTR(value));
                }

                public static IEnumerable<object[]> SuccessData
                {
                    get
                    {
                        yield return new object[] { "Date with zero time", new DateTime(2015, 5, 28), "5/28/2015" };
                        yield return new object[] { "Date with non-zero time", new DateTime(2015, 5, 28, 18, 54, 36), "5/28/2015 6:54:36 PM" };
                        yield return new object[] { "Zero date with non-zero time", VBScriptConstants.ZeroDate.Add(new TimeSpan(18, 54, 36)), "6:54:36 PM" };
                        yield return new object[] { "Zero date with zero time", VBScriptConstants.ZeroDate, "12:00:00 AM" };
                    }
                }
            }
        }
    }
}
