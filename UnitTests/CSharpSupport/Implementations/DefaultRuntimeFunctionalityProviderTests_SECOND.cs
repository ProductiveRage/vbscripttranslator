using System;
using System.Collections.Generic;
using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        /// <summary>
        /// This is EXTREMELY close to CDATE.. but not exactly the same (the Null case is the only difference, I think - that CDATE throws an error while this returns Null)
        /// </summary>
        public class SECOND
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().SECOND(value));
            }
            
            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().SECOND(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {

                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().SECOND(value);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object value)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().SECOND(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, 0 };
                    yield return new object[] { "Null", DBNull.Value, DBNull.Value };
                    yield return new object[] { "Zero", null, 0 };
                    yield return new object[] { "Minus one", -1, 0 };
                    yield return new object[] { "Minus 400", -400, 0 };
                    yield return new object[] { "Minus 400.2", -400.2, 0 };
                    yield return new object[] { "Minus 400.008", -400.008, 31 };
                    yield return new object[] { "Minus 400.8", -400.8, 0 };
                    yield return new object[] { "Plus 400.2", 400.2, 0 };
                    yield return new object[] { "Plus 400.8", 400.8, 0 };
                    yield return new object[] { "Plus 40000.001", 40000.001, 26 };
                    yield return new object[] { "Plus 40000.01", 40000.01, 24 };
                    yield return new object[] { "Plus 40000.02", 40000.02, 48 };
                    yield return new object[] { "Plus 40000.08", 40000.08, 12 };
                    yield return new object[] { "String \"-400.008\"", "-400.008", 31 };
                    yield return new object[] { "String \"-400.8\"", "-400.8", 0 };
                    yield return new object[] { "String \"2009-10-11\"", "2009-10-11", 0 };
                    yield return new object[] { "String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44", 44 };
                    yield return new object[] { "A Date", new DateTime(2009, 7, 6, 20, 12, 44), 44 };

                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype(), 0 };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, DBNull.Value };
                    yield return new object[] { "Object with default property which is Zero", new exampledefaultpropertytype { result = 0 }, 0 };
                    yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, 44 };

                    // Some bizarre behaviour occurs at the very top end of the supported range - at the very last integer, when 0.02 is added the number of seconds is inconsistent
                    // with 0.02 being added to ANY (positive?) integer smaller than it; it changes from always being 53 to being 52 at the very last change.
                    // Some bizarre behaviour occurs at the very top end of the supported range - at the very last integer, when 0.002 is present as the time component, then the number
                    // of seconds is inconsistent with ANY other value in the acceptable range that has a 0.002 time component; it changes from always being 53 to being 52 at the very
                    // last chance.
                    yield return new object[] { "Minus 400.002", -400.002, 53 };
                    yield return new object[] { "Plus 2000000.002 (approx 2/3 of largest possible positive integer)", 2000000.002, 53 };
                    yield return new object[] { "One before the largest positive integer before overflow + 0.02", 2958464.002, 53 };
                    yield return new object[] { "Largest positive integer before overflow + 0.02", 2958465.002, 52 };
                    yield return new object[] { "Most negative possible value with .002 time component", -657434.002, 53 };

                    yield return new object[] { "Minus 400.9", -400.9, 0 };
                    yield return new object[] { "Plus 2000000.9 (approx 2/3 of largest possible positive integer)", 2000000.9, 0 };
                    yield return new object[] { "One before the largest positive integer before overflow + 0.9", 2958464.9, 0 };
                    yield return new object[] { "Largest positive integer before overflow + 0.9", 2958465.9, 59 };
                    yield return new object[] { "Most negative possible value with .9 time component", -657434.9, 0 };

                    // Overflow edge checks
                    yield return new object[] { "Largest positive integer before overflow", 2958465, 0 };
                    yield return new object[] { "Largest positive integer before overflow + 0.9", 2958465.9, 59 };
                    yield return new object[] { "Largest positive integer before overflow + 0.99", 2958465.99, 36 };
                    yield return new object[] { "Largest positive integer before overflow + 0.999", 2958465.999, 33 };
                    yield return new object[] { "Largest positive integer before overflow + 0.9999", 2958465.9999, 51 };
                    yield return new object[] { "Largest positive integer before overflow + 0.99999", 2958465.99999, 59 };
                    yield return new object[] { "Largest positive integer before overflow + 0.999999999", 2958465.999999999, 59 }; // This is the most 9s that must be allowed before an overflow occurs

                    yield return new object[] { "Largest negative integer before overflow", -657434, 0 };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Blank string", ""};
                    yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" } };
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

            public static IEnumerable<object[]> OverflowData
            {
                get
                {
                    yield return new object[] { "Large number (12388888888888.2)", 12388888888888.2 };
                    yield return new object[] { "Object with default property which is a large number (12388888888888.2)", new exampledefaultpropertytype { result = 12388888888888.2 } };
                    
                    yield return new object[] { "Smallest positive integer that overflows", 2958466 };
                    yield return new object[] { "Smallest negative integer that overflows", -657435 };

                    yield return new object[] { "Largest positive integer before overflow + 0.9999999999", 2958465.9999999999 }; // This is the fewest 9s that must be allowed before an overflow occurs
                }
            }
        }
    }
}
