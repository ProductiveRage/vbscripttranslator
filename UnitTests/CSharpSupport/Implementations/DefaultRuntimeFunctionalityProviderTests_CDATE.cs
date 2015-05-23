using System;
using System.Collections.Generic;
using System.Linq;
using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CDATE
        {
            /// <summary>
            /// Values that go through a CDATE / CDBL cycle should come out unchanged. There are some unfortunate exceptions to this in the current implementation, special handling
            /// had to be added around the integer that represents the greatest possible date in VBScript (2958465 / 9999-12-31) and so any non-integer values based upon that number
            /// will fail a round trip by a small margin (eg. 2958465.9), but other values should pass through a round trip unaltered (to a certain level of precision).
            /// </summary>
            [Fact]
            public void RoundTripConversionCases()
            {
                var _ = DefaultRuntimeSupportClassFactory.Get();
                var values = new[] { 0, 1, -1, -400, -400.2, -400.008, -400.8, 400.2, 400.8, 40000.001, 40000.01, 40000.02, 40000.08, -400.002, 2000000.002, 2958464.002, -657434.002, -400.9, 2000000.9, 2958464.9, -657434.9, -657434 };
                Assert.Equal(
                    values,
                    values.Select(value => _.CDBL(_.CDATE(value))).ToArray()
                );
            }

            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object value, DateTime expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().CDATE(value));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDATE(value);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDATE(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDATE(value);
                });
            }

            [Theory, MemberData("OverflowData")]
            public void OverflowCases(string description, object value)
            {
                Assert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CDATE(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Empty", null, new DateTime(1899, 12, 30, 0, 0, 0) };
                    yield return new object[] { "Zero", null, new DateTime(1899, 12, 30, 0, 0, 0) };
                    yield return new object[] { "Minus one", -1, new DateTime(1899, 12, 29, 0, 0, 0) };
                    yield return new object[] { "Minus 400", -400, new DateTime(1898, 11, 25, 0, 0, 0) };
                    yield return new object[] { "Minus 400.2", -400.2, new DateTime(1898, 11, 25, 4, 48, 0) }; // This is nuts! It's like -400 then +0.2, but that's what VBScript seems to do..
                    yield return new object[] { "Minus 400.8", -400.8, new DateTime(1898, 11, 25, 19, 12, 0) };
                    yield return new object[] { "Plus 40000", 40000, new DateTime(2009, 7, 6, 0, 0, 0) };
                    yield return new object[] { "Plus 40000.2", 40000.2, new DateTime(2009, 7, 6, 4, 48, 0) };
                    yield return new object[] { "Plus 40000.8", 40000.8, new DateTime(2009, 7, 6, 19, 12, 0) };
                    yield return new object[] { "String \"-400.2\"", "-400.2", new DateTime(1898, 11, 25, 4, 48, 0) };
                    yield return new object[] { "String \"40000.2\"", "40000.2", new DateTime(2009, 7, 6, 4, 48, 0) };
                    yield return new object[] { "String \"2009-10-11\"", "2009-10-11", new DateTime(2009, 10, 11) };
                    yield return new object[] { "String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44", new DateTime(2009, 10, 11, 20, 12, 44) };
                    yield return new object[] { "A Date", new DateTime(2009, 7, 6, 20, 12, 44), new DateTime(2009, 7, 6, 20, 12, 44) };
                    
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype(), new DateTime(1899, 12, 30, 0, 0, 0) };
                    yield return new object[] { "Object with default property which is Zero", new exampledefaultpropertytype(), new DateTime(1899, 12, 30, 0, 0, 0) };
                    yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, new DateTime(2009, 10, 11, 20, 12, 44) };

                    // These may go some way to explaining the -400.2 case above, it appears that the negative sign is removed from fractional values in VBScript
                    yield return new object[] { "Plus 0.1", 0.1, new DateTime(1899, 12, 30, 2, 24, 0) };
                    yield return new object[] { "Plus 0.2", 0.2, new DateTime(1899, 12, 30, 4, 48, 0) };
                    yield return new object[] { "Plus 0.3", 0.3, new DateTime(1899, 12, 30, 7, 12, 0) };
                    yield return new object[] { "Plus 0.4", 0.4, new DateTime(1899, 12, 30, 9, 36, 0) };
                    yield return new object[] { "Plus 0.8", 0.8, new DateTime(1899, 12, 30, 19, 12, 0) };
                    yield return new object[] { "Plus 0.99", 0.99, new DateTime(1899, 12, 30, 23, 45, 36) };
                    yield return new object[] { "Minus 0.1", -0.1, new DateTime(1899, 12, 30, 2, 24, 0) };
                    yield return new object[] { "Minus 0.2", -0.2, new DateTime(1899, 12, 30, 4, 48, 0) };
                    yield return new object[] { "Minus 0.3", -0.3, new DateTime(1899, 12, 30, 7, 12, 0) };
                    yield return new object[] { "Minus 0.4", -0.4, new DateTime(1899, 12, 30, 9, 36, 0) };
                    yield return new object[] { "Minus 0.8", -0.8, new DateTime(1899, 12, 30, 19, 12, 0) };
                    yield return new object[] { "Minus 0.99", -0.99, new DateTime(1899, 12, 30, 23, 45, 36) };

                    // Overflow edge checks
                    yield return new object[] { "Largest positive integer before overflow", 2958465, new DateTime(9999, 12, 31) };
                    yield return new object[] { "Largest negative integer before overflow", -657434, new DateTime(100, 1, 1) };
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
