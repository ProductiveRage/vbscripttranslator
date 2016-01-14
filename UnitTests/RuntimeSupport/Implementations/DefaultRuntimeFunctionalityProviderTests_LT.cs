using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using System;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
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
                    DefaultRuntimeSupportClassFactory.Get().LT(null, null)
                );
            }

            [Fact]
            public void NullComparedToNullIsNull()
            {
                Assert.Equal(
                    DBNull.Value,
                    DefaultRuntimeSupportClassFactory.Get().LT(DBNull.Value, DBNull.Value)
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
                    DefaultRuntimeSupportClassFactory.Get().LT(DBNull.Value, nothing);
                });
            }

            [Fact]
            public void NothingComparedToNothingErrors()
            {
                var nothing = VBScriptConstants.Nothing;
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().LT(nothing, nothing);
                });
            }

            // Empty appears to be treated as zero
            [Fact]
            public void ZeroIsNotLessThanEmpty()
            {
                Assert.Equal(
                    false,
                    DefaultRuntimeSupportClassFactory.Get().LT(0, null)
                );
            }
            [Fact]
            public void EmptyIsNotLessThanZero()
            {
                Assert.Equal(
                    false,
                    DefaultRuntimeSupportClassFactory.Get().LT(null, 0)
                );
            }
            [Fact]
            public void MinusOneIsLessThanEmpty()
            {
                Assert.Equal(
                    true,
                    DefaultRuntimeSupportClassFactory.Get().LT(-1, null)
                );
            }
            [Fact]
            public void EmptyIsNotLessThanMinusOne()
            {
                Assert.Equal(
                    false,
                    DefaultRuntimeSupportClassFactory.Get().LT(null, -1)
                );
            }
            [Fact]
            public void EmptyIsLessThanPlusOne()
            {
                Assert.Equal(
                    true,
                    DefaultRuntimeSupportClassFactory.Get().LT(null, 1)
                );
            }
            [Fact]
            public void PlusOneIsNotLessThanEmpty()
            {
                Assert.Equal(
                    false,
                    DefaultRuntimeSupportClassFactory.Get().LT(1, null)
                );
            }
        }
    }
}
