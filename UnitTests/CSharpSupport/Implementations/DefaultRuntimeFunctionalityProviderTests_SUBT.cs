using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class SUBT
        {
            public class SingleArgument
            {
                [Theory, MemberData("SuccessData")]
                public void SuccessCases(string description, object value, object expectedResult)
                {
                    Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().SUBT(value));
                }

                [Theory, MemberData("TypeMismatchData")]
                public void TypeMismatchCases(string description, object value)
                {
                    Assert.Throws<TypeMismatchException>(() =>
                    {
                        DefaultRuntimeSupportClassFactory.Get().SUBT(value);
                    });
                }

                [Theory, MemberData("ObjectVariableNotSetData")]
                public void ObjectVariableNotSetCases(string description, object value)
                {
                    Assert.Throws<ObjectVariableNotSetException>(() =>
                    {
                        DefaultRuntimeSupportClassFactory.Get().SUBT(value);
                    });
                }

                [Theory, MemberData("ObjectDoesNotSupportPropertyOrMemberData")]
                public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object value)
                {
                    Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
                    {
                        DefaultRuntimeSupportClassFactory.Get().SUBT(value);
                    });
                }

                public static IEnumerable<object[]> SuccessData
                {
                    get
                    {
                        yield return new object[] { "Empty", null, (Int16)0 };
                        yield return new object[] { "Null", DBNull.Value, DBNull.Value };
                        yield return new object[] { "Byte zero", (byte)0, (byte)0 };
                        yield return new object[] { "Integer zero", (Int16)0, (Int16)0 };

                        yield return new object[] { "CByte(1)", (byte)1, (Int16)(-1) }; // Non-zero Byte values have to change type since Byte can't represent negative values
                        yield return new object[] { "CInt(1)", (Int16)1, (Int16)(-1) };
                        yield return new object[] { "CLng(1)", 1, -1 };
                        yield return new object[] { "CDate(1)", VBScriptConstants.ZeroDate.AddDays(1), VBScriptConstants.ZeroDate.AddDays(-1) };
                        yield return new object[] { "CCur(1)", 1m, -1m };
                        yield return new object[] { "CDbl(1)", 1d, -1d };

                        // Strings must be numeric (and if so, they'll be parsed as Doubles)
                        yield return new object[] { "String \"1\"", "1", -1d };
                        yield return new object[] { "String \"-1\"", "-1", 1d };

                        // These are cases which type type changes to avoid an overflow
                        yield return new object[] { "CInt(-32768)", (Int16)(-32768), 32768 };
                        yield return new object[] { "CLng(-2147483648)", -2147483648, 2147483648d };
                        yield return new object[] { "max-date", VBScriptConstants.LatestPossibleDate, -VBScriptConstants.LatestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays };

                        // The Currency type has identical min and max values (just with a flipped sign) so there is no edge value that overflows when negated
                        yield return new object[] { "max-currency-value", VBScriptConstants.MaxCurrencyValue, VBScriptConstants.MinCurrencyValue };
                        yield return new object[] { "min-currency-value", VBScriptConstants.MinCurrencyValue, VBScriptConstants.MaxCurrencyValue };
                    }
                }

                public static IEnumerable<object[]> TypeMismatchData
                {
                    get
                    {
                        yield return new object[] { "True", true };
                        yield return new object[] { "False", false };
                        yield return new object[] { "Blank string", "" };
                        yield return new object[] { "String \"a\"", "a" };
                    }
                }

                public static IEnumerable<object[]> ObjectVariableNotSetData
                {
                    get
                    {
                        yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                    }
                }

                public static IEnumerable<object[]> ObjectDoesNotSupportPropertyOrMemberData
                {
                    get
                    {
                        yield return new object[] { "Object without default member", new Object() };
                    }
                }
            }
        }
    }
}
