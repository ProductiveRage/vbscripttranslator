using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CSharpSupport;
using CSharpSupport.Attributes;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class TYPENAME
        {
            [Theory, MemberData("Data")]
            public void Cases(string description, object value, string expectedTypeName)
            {
                Assert.Equal(expectedTypeName, DefaultRuntimeSupportClassFactory.Get().TYPENAME(value));
            }

            public static IEnumerable<object[]> Data
            {
                get
                {
                    yield return new object[] { "Empty", null, "Empty" };
                    yield return new object[] { "Null", DBNull.Value, "Null" };
                    yield return new object[] { "Nothing", VBScriptConstants.Nothing, "Nothing" };
                    yield return new object[] { "True", true, "Boolean" };
                    yield return new object[] { "False", false, "Boolean" };
                    yield return new object[] { "Byte", (Byte)1, "Byte" };
                    yield return new object[] { "VBScript Integer (Int16)", (Int16)1, "Integer" };
                    yield return new object[] { "VBScript Long (Int32)", (Int32)1, "Long" };
                    yield return new object[] { "VBScript Double", (double)1, "Double" };
                    yield return new object[] { "VBScript Currency (Decimal)", (decimal)1, "Currency" };
                    yield return new object[] { "Date", new DateTime(2015, 5, 18, 20, 35, 0), "Date" };
                    yield return new object[] { "Date without time component", new DateTime(2015, 5, 18), "Date" };
                    yield return new object[] { "VBScript time (ZeroDate with time component)", VBScriptConstants.ZeroDate.Add(new TimeSpan(20, 35, 0)), "Date" };
                    yield return new object[] { "Scripting Dictionary", Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary")), "Dictionary" };
                    yield return new object[] { "Translated Class", new exampledefaultpropertytype(), "ExampleDefaultPropertyType" };
                    yield return new object[] { "COM Visible Class", new ComVisibleClass(), "ComVisibleClass" };
                    yield return new object[] { "Non-COM-Visible Class derived from a COM Visible Class", new NonComVisibleClassDerivedFromComVisibleClass(), "ComVisibleClass" };
                    yield return new object[] { "Non-COM-Visible Class", new NonComVisibleClass(), "Object" };
                }
            }

            [ComVisible(true)]
            private class ComVisibleClass { }

            private class NonComVisibleClassDerivedFromComVisibleClass : ComVisibleClass { }

            private class NonComVisibleClass { }

            /// <summary>
            /// This is an example of the type of class that may be emitted by the translation process, one with a parameter-less default member
            /// </summary>
            [SourceClassName("ExampleDefaultPropertyType")]
            private class exampledefaultpropertytype
            {
                [IsDefault]
                public object result { get; set; }
            }
        }
    }
}
