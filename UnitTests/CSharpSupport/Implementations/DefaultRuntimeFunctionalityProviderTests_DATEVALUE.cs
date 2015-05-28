using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class DATEVALUE
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, DateTime expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().DATEVALUE(value));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEVALUE(value);
                });
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEVALUE(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEVALUE(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "String \"2009-10-11\"", "2009-10-11", new DateTime(2009, 10, 11) };
                    yield return new object[] { "String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44", new DateTime(2009, 10, 11) };
                    yield return new object[] { "A Date", new DateTime(2009, 7, 6, 20, 12, 44), new DateTime(2009, 7, 6) };
                    yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, new DateTime(2009, 10, 11) };

                    // Note: We could go to town with test cases for the various string formats that VBScript supports, but the DATEVALUE implementation backs onto the DateParser and
                    // it would be duplication of effort going through everything again here (plus we'd need a way to set the default year for two segment "dynamic year" date strings,
                    // such as "1 5" (which could be the 1st of May in the current year or the 5th of January, depending upon culture)
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Empty", null };
                    yield return new object[] { "Zero", null };
                    yield return new object[] { "Minus one", -1 };
                    yield return new object[] { "Minus 400", -400 };
                    yield return new object[] { "Plus 40000", 40000 };
                    yield return new object[] { "Blank string", "" };
                    yield return new object[] { "String \"-400.2\"", "-400.2" };
                    yield return new object[] { "String \"40000.2\"", "40000.2" };
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype() };
                    yield return new object[] { "Object with default property which is Zero", new exampledefaultpropertytype { result = 0 } };
                    yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" } };
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

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
                }
            }
        }
    }
}
