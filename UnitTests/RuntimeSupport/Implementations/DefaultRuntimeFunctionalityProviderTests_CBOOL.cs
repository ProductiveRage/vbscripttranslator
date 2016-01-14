using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Attributes;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class CBOOL
        {
            [Theory, MemberData("TrueData")]
            public void TrueCases(string description, object value)
            {
                Assert.True(DefaultRuntimeSupportClassFactory.Get().CBOOL(value));
            }

            [Theory, MemberData("FalseData")]
            public void FalseCases(string description, object value)
            {
                Assert.False(DefaultRuntimeSupportClassFactory.Get().CBOOL(value));
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CBOOL(value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CBOOL(value);
                });
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().CBOOL(value);
                });
            }

            public static IEnumerable<object[]> TrueData
            {
                get
                {
                    yield return new object[] { "Number 1", 1 };
                    yield return new object[] { "Number 2", 2 };
                    yield return new object[] { "Number -1", -1 };
                    yield return new object[] { "Number -2", -2 };
                    yield return new object[] { "String \"true\"", "true" };
                    yield return new object[] { "String \"True\"", "True" };
                    yield return new object[] { "String \"TRUE\"", "TRUE" };
                    yield return new object[] { "Boolean True", true };
                    yield return new object[] { "Date other than zero", VBScriptConstants.ZeroDate.AddSeconds(1) };
                    yield return new object[] { "Object with default property which is string \"true\"", new exampledefaultpropertytype { result = "true" } };
                }
            }

            public static IEnumerable<object[]> FalseData
            {
                get
                {
                    yield return new object[] { "Empty", null };
                    yield return new object[] { "Zero", 0 };
                    yield return new object[] { "String \"false\"", "false" };
                    yield return new object[] { "String \"False\"", "False" };
                    yield return new object[] { "String \"FALSE\"", "FALSE" };
                    yield return new object[] { "Boolean False", false };
                    yield return new object[] { "Date Zero", VBScriptConstants.ZeroDate };
                    yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype() };
                }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null", DBNull.Value };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Blank String", "" };
                    yield return new object[] { "Whitespace", " " };
                    yield return new object[] { "Unintialised array", new object[0] };
                    yield return new object[] { "String \"FALSE \"", "FALSE " };
                }
            }
        }
    }
}
