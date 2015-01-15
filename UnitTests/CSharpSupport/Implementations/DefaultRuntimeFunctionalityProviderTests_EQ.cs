using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class EQ
        {
            [Fact]
            public void MinusOneDoesNotEqualEmpty()
            {
                // Non-zero numeric values compared to Empty for equality always return false
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(1, null)
                );
            }

            [Fact]
            public void PlusOneDoesNotEqualEmpty()
            {
                // Non-zero numeric values compared to Empty for equality always return false
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1, null)
                );
            }

            [Fact]
            public void ZeroEqualsEmpty()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(0, null)
                );
            }

            [Fact]
            public void MinusOneComparedToNullIsNull()
            {
                // Non-zero numeric values compared to Empty for equality always return false
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(1, DBNull.Value)
                );
            }

            [Fact]
            public void PlusOneComparedToNullIsNull()
            {
                // Non-zero numeric values compared to Empty for equality always return false
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1, DBNull.Value)
                );
            }

            [Fact]
            public void ZeroComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(0, DBNull.Value)
                );
            }

            [Fact]
            public void MinusOneEqualsTrue()
            {
                // -1 and True are considered to be the same, as are 0 and False
                // - No other numbers are considered to be equals of booleans (not -1.1, not -2, not 1, not 2)
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1, true)
                );
            }

            [Fact]
            public void ZeroEqualsFalse()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(0, false)
                );
            }

            [Fact]
            public void MinusOnePointOneDoesNotEqualTrue()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1, true)
                );
            }

            [Fact]
            public void MinusOnePointOneDoesNotEqualFalse()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1, false)
                );
            }

            [Fact]
            public void PlusOneDoesNotEqualTrue()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(1, true)
                );
            }

            [Fact]
            public void PlusOneDoesNotEqualFalse()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(1, false)
                );
            }

            [Fact]
            public void EmptyStringEqualsEmpty()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ("", null)
                );
            }

            [Fact]
            public void EmptyStringComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ("", DBNull.Value)
                );
            }

            [Fact]
            public void WhiteSpaceStringDoesNotEqualEmpty()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(" ", null)
                );
            }

            [Fact]
            public void WhiteSpaceStringComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(" ", DBNull.Value)
                );
            }

            // TODO: Strings vs booleans

            // TODO: Strings vs numbers

            [Fact]
            public void TrueDoesNotEqualEmpty()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(true, null)
                );
            }

            [Fact]
            public void FalseDoesNotEqualEmpty()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(false, null)
                );
            }

            [Fact]
            public void TrueComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(true, DBNull.Value)
                );
            }

            [Fact]
            public void FalseComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(false, DBNull.Value)
                );
            }

            // TODO: Objects - with and without default members
            // - ADODB.Field?
            // - Does VBScript consider "Item" or "Value" as default even if not DispId zero? Is this only for non-VBScript-originating types>
            // ********** Leave all of this to the VAL tests

            [Fact]
            public void DefaultMemberIsRequiredForPropertyReferences()
            {
                Assert.Throws<ArgumentException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().EQ(new object(), "Test");
                });
            }

            [Fact]
            public void IDispatchDefaultMemberIsAccessedOnPropertyReference()
            {
                dynamic recordset = Activator.CreateInstance(Type.GetTypeFromProgID("ADODB.Recordset"));
                recordset.Fields.Append("Name", 200, 50, 4);
                recordset.Open();
                recordset.AddNew();
                recordset.Fields["Name"].Value = "Test";
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(recordset.Fields["Name"], "Test")
                );
            }

            [Fact]
            public void ImplicitNamedDefaultMemberIsAccessedOnPropertyReference()
            {
                dynamic recordset = Activator.CreateInstance(Type.GetTypeFromProgID("ADODB.Recordset"));
                recordset.Fields.Append("Name", 200, 50, 4);
                recordset.Open();
                recordset.AddNew();
                recordset.Fields["Name"].Value = "Test";
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(recordset.Fields["Name"], "Test")
                );
            }
        }
    }
}
