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
        public class HOUR
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().HOUR(value));
            }
            
            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().HOUR(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetExceptionData")]
            public void ObjectVariableNotSetExceptionCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().HOUR(value);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object value)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().HOUR(value);
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
                    yield return new object[] { "Minus 400.2", -400.2, 4 };
                    yield return new object[] { "Minus 400.8", -400.8, 19 };
                    yield return new object[] { "Plus 400.2", 400.2, 4 };
                    yield return new object[] { "Plus 400.8", 400.8, 19 };
                    yield return new object[] { "Plus 40000", 40000, 0 };
                    yield return new object[] { "String \"-400.2\"", "-400.2", 4 };
                    yield return new object[] { "String \"2009-10-11\"", "2009-10-11", 0 };
                    yield return new object[] { "String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44", 20 };
                    yield return new object[] { "A Date", new DateTime(2009, 7, 6, 20, 12, 44), 20 };
                    
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype(), 0 };
                    yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, DBNull.Value };
                    yield return new object[] { "Object with default property which is Zero", new exampledefaultpropertytype(), 0 };
                    yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, 20 };

                    // Overflow edge checks
                    yield return new object[] { "Largest positive integer before overflow", 2958465, 0 };
                    yield return new object[] { "Largest positive integer before overflow + 0.9", 2958465.9, 21 };
                    yield return new object[] { "Largest positive integer before overflow + 0.99", 2958465.99, 23 };
                    yield return new object[] { "Largest positive integer before overflow + 0.999", 2958465.999, 23 };
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

            public static IEnumerable<object[]> ObjectVariableNotSetExceptionData
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

            /// <summary>
            /// This is an example of the type of class that may be emitted by the translation process, one with a parameter-less default member
            /// </summary>
            [TranslatedProperty("ExampleDefaultPropertyType")]
            private class exampledefaultpropertytype
            {
                [IsDefault]
                public object result { get; set; }
            }
        }
    }
}
