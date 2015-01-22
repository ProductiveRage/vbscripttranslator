using CSharpSupport;
using CSharpSupport.Exceptions;
using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        // Note: As with the EQ tests, there won't be a lot of cases here around comparing values extracted from object references since that logic is dealt with
        // by the VAL method (and once a non-object-reference value has been obtained, the same logic as illustrated below will be followed)
        public class LT
        {
            [Fact]
            public void EmptyIsNotLessThanEmpty()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().LT(null, null)
                );
            }

            [Fact]
            public void NullComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    GetDefaultRuntimeFunctionalityProvider().LT(DBNull.Value, DBNull.Value)
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
                    GetDefaultRuntimeFunctionalityProvider().LT(DBNull.Value, nothing);
                });
            }

            [Fact]
            public void NothingComparedToNothingErrors()
            {
                var nothing = VBScriptConstants.Nothing;
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    GetDefaultRuntimeFunctionalityProvider().LT(nothing, nothing);
                });
            }

            // Empty appears to be treated as zero
            [Fact]
            public void ZeroIsNotLessThanEmpty()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().LT(0, null)
                );
            }
            [Fact]
            public void EmptyIsNotLessThanZero()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().LT(null, 0)
                );
            }
            [Fact]
            public void MinusOneIsLessThanEmpty()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().LT(-1, null)
                );
            }
            [Fact]
            public void EmptyIsNotLessThanMinusOne()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().LT(null, -1)
                );
            }
            [Fact]
            public void EmptyIsLessThanPlusOne()
            {
                Assert.Equal(
                    true,
                    GetDefaultRuntimeFunctionalityProvider().LT(null, 1)
                );
            }
            [Fact]
            public void PlusOneIsNotLessThanEmpty()
            {
                Assert.Equal(
                    false,
                    GetDefaultRuntimeFunctionalityProvider().LT(1, null)
                );
            }
        }
    }
}
