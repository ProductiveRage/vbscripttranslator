using CSharpSupport;
using CSharpSupport.Exceptions;
using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        // Note: There are a class of testst that are not present here - where one or both sides of the comparison are an object reference. In these cases, the EQ
        // implementation on the DefaultRuntimeFunctionalityProvider class pushes these through the VAL method in order to extract a value for comparison (if this
        // fails then a Type Mismatch error is raised). When values are present on both sides, the logic tested here is applied. The tests for the VAL method will
        // cover the logic regarding this, we don't need to duplicate it here. The same goes for arrays - the VAL logic will handle it.
        public class EQ
        {
            [Fact]
            public void EmptyEqualsEmpty()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(null, null)
                );
            }

            [Fact]
            public void NullComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().EQ(DBNull.Value, DBNull.Value)
                );
            }

            /// <summary>
            /// Anything compared to Nothing will error, this is just an example case to illustrate that (if ANYTHING would get a free pass it would be DBNull.Value
            /// but not even it does)
            /// </summary>
            [Fact]
            public void NullComparedToNothingErrors()
            {
                var nothing = VBScriptConstants.Nothing;
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().EQ(DBNull.Value, nothing);
                });
            }

            [Fact]
            public void NothingComparedToNothingErrors()
            {
                var nothing = VBScriptConstants.Nothing;
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().EQ(nothing, nothing);
                });
            }

            [Fact]
            public void MinusOneDoesNotEqualEmpty()
            {
                // Non-zero numeric values compared to Empty for equality always return false
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1, null)
                );
            }

            [Fact]
            public void PlusOneDoesNotEqualEmpty()
            {
                // Non-zero numeric values compared to Empty for equality always return false
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(1, null)
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
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1.1, true)
                );
            }

            [Fact]
            public void MinusOnePointOneDoesNotEqualFalse()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(-1.1, false)
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
            public void EmptyStringDoesNotEqualsTrue()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ("", true)
                );
            }

            [Fact]
            public void EmptyStringDoesNotEqualsFalse()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ("", false)
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

            [Fact]
            public void NumericContentStringValueDoesNotEqualNumericValue()
            {
                // Recall that the VBScript expression ("12" = 12) will return true, but if v12String = "12" and v12 = 12 then (v12String = v12) will return
                // false. For cases where string or number literals are present in the comparison, the translator must cast the other side so that they both
                // are consistent but the EQ method does not have to deal with it - so, here, "12" does not equal 12.
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ("12", 12)
                );
            }

            [Fact]
            public void BooleanContentStringValueDoesNotEqualBooleanValue()
            {
                // See the note in NumericContentStringValueDoesNotEqualNumericValue about literals - the same applies here; while ("True" = True) will return
                // true, if vTrueString = "True" and vTrue = True then (vTrueString = vTrue) return false and it is only this latter case that EQ must deal
                // with, any special handling regarding literals must be dealt with by the translator before getting to EQ.
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ("True", true)
                );
            }

            [Fact]
            public void TrueDoesNotEqualEmpty()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().EQ(true, null)
                );
            }

            [Fact]
            public void FalseEqualsEmpty()
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

            [Fact]
            public void TrueEqualsMinusOne()
            {
                // Dim vTrue, vMinusOne: vTrue = True: vMinusOne = -1: If (vTrue = vMinusOne) Then ' True
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(true, -1)
                );
            }

            [Fact]
            public void TrueEqualsDoubleMinusOne()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(true, -1.0d)
                );
            }

            [Fact]
            public void FalseEqualsZero()
            {
                // Dim vFalse, vZero: vFalse = False: vZero = 0: If (vFalse = vZero) Then ' True
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().EQ(false, 0)
                );
            }
        }
    }
}
