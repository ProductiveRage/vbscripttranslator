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
        public class MONTH
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().MONTH(value));
            }
            
            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MONTH(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MONTH(value);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object value)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().MONTH(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, 12 };
                    yield return new object[] { "Null", DBNull.Value, DBNull.Value };
                    yield return new object[] { "Zero", null, 12 };
                    yield return new object[] { "Minus one", -1, 12 };
                    yield return new object[] { "Minus 400", -400, 11 };
                    yield return new object[] { "Plus 40000", 40000, 7 };
                    yield return new object[] { "String \"-400.2\"", "-400.2", 11 };
                    yield return new object[] { "String \"40000.2\"", "40000.2", 7 };
                    yield return new object[] { "String \"2009-10-11\"", "2009-10-11", 10 };
                    yield return new object[] { "String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44", 10 };
                    yield return new object[] { "A Date", new DateTime(2009, 7, 6, 20, 12, 44), 7 };
                    
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype(), 12 };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, DBNull.Value };
                    yield return new object[] { "Object with default property which is Zero", new exampledefaultpropertytype { result = 0 }, 12 };
                    yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, 10 };

                    // Overflow edge checks
                    yield return new object[] { "Largest positive integer before overflow", 2958465, 12 };
                    yield return new object[] { "Largest negative integer before overflow", -657434, 1 };
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
                }
            }
        }
    }
}
