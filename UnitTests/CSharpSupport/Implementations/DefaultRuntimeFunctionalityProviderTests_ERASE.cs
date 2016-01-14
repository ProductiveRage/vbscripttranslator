using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class ERASE
        {
            public class ByRefMethodSignature
            {
                [Fact]
                public void ArrayTargetShouldBeReplacedWithEmptyArray()
                {
                    object target = new object[] { 123 };
                    DefaultRuntimeSupportClassFactory.Get().ERASE(target, erasedTarget => { target = erasedTarget; });
                    Assert.Equal(new object[0], target);
                }

                [Theory, MemberData("TypeMismatchData")]
                public void TypeMismatchCases(string description, object target)
                {
                    Assert.Throws<TypeMismatchException>(() =>
                    {
                        DefaultRuntimeSupportClassFactory.Get().ERASE(target, erasedTarget => { target = erasedTarget; });
                    });
                }

                public static IEnumerable<object[]> TypeMismatchData
                {
                    get
                    {
                        yield return new object[] { "Empty", null };
                        yield return new object[] { "Null", DBNull.Value };
                        yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                        yield return new object[] { "Zero", 0 };
                        yield return new object[] { "Blank string", "" };
                        yield return new object[] { "A date", new DateTime(2009, 10, 11, 20, 12, 44) };
                        yield return new object[] { "Object with default property which is an array", new exampledefaultpropertytype { result = new object[] { 123 } } };
                    }
                }
            }

            public class IndirectArrayAccessMethodSignature
            {
                [Fact]
                public void NestedArrayTargetShouldBeReplacedWithEmptyArray()
                {
                    object target = new object[] { new object[] { 123 } };
                    DefaultRuntimeSupportClassFactory.Get().ERASE(target, 0);
                    Assert.Equal(new object[0], ((object[])target)[0]);
                }

                [Theory, MemberData("TypeMismatchData")]
                public void TypeMismatchCases(string description, object value, object[] arguments)
                {
                    Assert.Throws<TypeMismatchException>(() =>
                    {
                        DefaultRuntimeSupportClassFactory.Get().ERASE(value, arguments);
                    });
                }

                [Theory, MemberData("SubscriptOutOfRangeData")]
                public void SubscriptOutOfRangeCases(string description, object value, object[] arguments)
                {
                    Assert.Throws<SubscriptOutOfRangeException>(() =>
                    {
                        DefaultRuntimeSupportClassFactory.Get().ERASE(value, arguments);
                    });
                }

                public static IEnumerable<object[]> TypeMismatchData
                {
                    get
                    {
                        yield return new object[] { "Empty", null, new object[] { 1 } };
                        yield return new object[] { "Null", DBNull.Value, new object[] { 1 } };
                        yield return new object[] { "Nothing", VBScriptConstants.Nothing, new object[] { 1 } };
                        yield return new object[] { "Zero", 0, new object[] { 1 } };
                        yield return new object[] { "Blank string", "", new object[] { 1 } };
                        yield return new object[] { "A date", new DateTime(2009, 10, 11, 20, 12, 44), new object[] { 1 } };
                        yield return new object[] { "Object with default property which is an array", new exampledefaultpropertytype { result = new object[] { 123 } }, new object[] { 1 } };
                        yield return new object[] { "Array target where specified element is not an array", new object[] { 123 }, new object[] { 0 } };
                    }
                }

                public static IEnumerable<object[]> SubscriptOutOfRangeData
                {
                    get
                    {
                        yield return new object[] { "Array target with zero arguments", new object[] { 123 }, new object[0] };
                        yield return new object[] { "Array target with negative argument", new object[] { 123 }, new object[] { -1 } };
                        yield return new object[] { "Array target with out-of-range positive argument", new object[] { 123 }, new object[] { 1 } };
                        yield return new object[] { "Array target with too many dimensions", new object[] { 123 }, new object[] { 0, 0 } };
                    }
                }
            }
        }
    }
}
