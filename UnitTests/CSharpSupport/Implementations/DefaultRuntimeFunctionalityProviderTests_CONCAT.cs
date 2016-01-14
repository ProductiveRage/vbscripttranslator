using System;
using System.Collections.Generic;
using System.Globalization;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Attributes;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using VBScriptTranslator.UnitTests.Shared;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CONCAT  : CultureOverridingTests // The date-handling is culture-specific, so we need to explicitly specify what culture we're testing against
        {
            public CONCAT() : base(new CultureInfo("en-GB")) { }

            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object l, object r, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().CONCAT(l, r));
            }
            
            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object l, object r)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CONCAT(l, r);
                });
            }

            [Theory, MemberData("ObjectDoesNotSupportPropertyOrMemberData")]
            public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object l, object r)
            {
                Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CONCAT(l, r);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object l, object r)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CONCAT(l, r);
                });
            }

            [Theory, MemberData("OutOfStringSpaceData")]
            public void OutOfStringSpaceCases(string description, object l, object r)
            {
                Assert.Throws<OutOfStringSpaceException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CONCAT(l, r);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty & Empty = Blank string", null, null, "" };
                    yield return new object[] { "Null & Null = Null", DBNull.Value, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Empty & Null = Blank string", null, DBNull.Value, "" };
                    yield return new object[] { "Null & Empty = Blank string", DBNull.Value, null, "" };
                    
                    yield return new object[] { "\"a\" & Empty = \"a\"", "a", null, "a" };
                    yield return new object[] { "Empty & \"a\" = \"a\"", null, "a", "a" };
                    yield return new object[] { "\"a\" & Null = \"a\"", "a", DBNull.Value, "a" };
                    yield return new object[] { "Null & \"a\" = \"a\"", DBNull.Value, "a", "a" };

                    yield return new object[] { "\"a\" & Chr(0) = \"a\"", "a", (char)0, "a" + (char)0 };

                    yield return new object[] { "\"a\" & \"bc\" = \"abc\"", "a", "bc", "abc" }; // Best do at least one simple case! :)

                    yield return new object[] { "Blank string & 1 = \"1\"", "", 1, "1" };
                    yield return new object[] { "Blank string & (DateSerial(2015, 6, 18) + TimeSerial(22, 3, 56)) = \"18/06/2015 22:03:56\"", "", new DateTime(2015, 6, 18, 22, 3, 56), "18/06/2015 22:03:56" };

                    // Note: We can't actually test the limits of string length since .net will seemingly not handle strings as long as VBScript - hopefully this is not
                    // an edge case that is ever encountered in real life! :S
                    // - In VBScript, this is acceptable (it is the greatest length of string that is): String(1073741822, " ")
                    // - In C#, this is NOT acceptable (OutOfMemoryException): new string(' ', 1073741822)
                    //   The largest length appears to be (through trial-and-error) 1073741791 (explained properly by Jon Skeet in the Stack Overflow answer http://stackoverflow.com/a/5367331)
                    //yield return new object[] { "StringThatIsPreciselyHalfOfVBScriptMaximum & StringThatIsPreciselyHalfOfVBScriptMaximum", StringThatIsPreciselyHalfOfVBScriptMaximum, " ", StringThatIsPreciselyHalfOfVBScriptMaximum + " " };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Blank String & Nothing", "", VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing & object with default property which is Nothing (error on first argument takes precedence)", VBScriptConstants.Nothing, new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
                }
            }

            public static IEnumerable<object[]> ObjectDoesNotSupportPropertyOrMemberData
            {
                get
                {
                    yield return new object[] { "Blank String & object with no default member", "", new exampletranslatedclasswithnodefaultmember() };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Blank String & object with default property which is Nothing", "", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
                }
            }

            public static IEnumerable<object[]> OutOfStringSpaceData
            {
                get
                {
                    var stringThatIsPreciselyHalfOfVBScriptMaximum = new string(' ', 536870911);
                    yield return new object[] { "(StringThatIsPreciselyHalfOfVBScriptMaximum & \" \") & StringThatIsPreciselyHalfOfVBScriptMaximum", stringThatIsPreciselyHalfOfVBScriptMaximum + " ", stringThatIsPreciselyHalfOfVBScriptMaximum };
                }
            }

            [SourceClassName("ExampleTranslatedClassWithNoDefaultMember")]
            private class exampletranslatedclasswithnodefaultmember { }
        }
    }
}
