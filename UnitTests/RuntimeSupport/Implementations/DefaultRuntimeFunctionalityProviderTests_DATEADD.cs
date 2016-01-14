using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    public static partial class DefaultRuntimeFunctionalityProviderTests
    {
        public class DATEADD
        {
            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, object interval, object number, object value, object expectedResult)
            {
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().DATEADD(interval, number, value));
            }

            /// <summary>
            /// For some reason, writing these cases using the xUnit Theory attribute in the same format as the other success cases failed, the .999999999 double values were
            /// getting rounded and precision was being lost before the DATEADD method was even being called! I'm not sure why, but I'm going to just write these separately
            /// rather than worrying about it too much.
            /// </summary>
            [Theory, MemberData("PrecisionEdgeCaseData")]
            public void PrecisionEdgeCases(string description, object interval, int numberBaseValue, int numberNumberOfNines, object value, object expectedResult)
            {
                var number = Convert.ToDouble(numberBaseValue.ToString() + "." + new string('9', numberNumberOfNines));
                Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().DATEADD(interval, number, value));
            }

            [Theory, MemberData("TypeMismatchData")]
            public void TypeMismatchCases(string description, object interval, object number, object value)
            {
                Assert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEADD(interval, number, value);
                });
            }

            [Theory, MemberData("InvalidProcedureCallOrArgumentData")]
            public void InvalidProcedureCallOrArgumentCases(string description, object interval, object number, object value)
            {
                Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEADD(interval, number, value);
                });
            }

            [Theory, MemberData("InvalidUseOfNullData")]
            public void InvalidUseOfNullCases(string description, object interval, object number, object value)
            {
                Assert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEADD(interval, number, value);
                });
            }

            [Theory, MemberData("ObjectVariableNotSetData")]
            public void ObjectVariableNotSetCases(string description, object interval, object number, object value)
            {
                Assert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Get().DATEADD(interval, number, value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    // Note: We could go to town with test cases for the various string formats that VBScript supports, but the DATEADD implementation backs onto the DateParser and
                    // it would be duplication of effort going through everything again here (plus we'd need a way to set the default year for two segment "dynamic year" date strings,
                    // such as "1 5" (which could be the 1st of May in the current year or the 5th of January, depending upon culture)

                    // Middle-of-the-road cases
                    yield return new object[] { "+1 second", "s", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 35) };
                    yield return new object[] { "+0 seconds", "s", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 second", "s", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 33) };
                    yield return new object[] { "+1 second (upper case interval)", "S", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 35) };
                    yield return new object[] { "+0 seconds (upper case interval)", "S", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 second (upper case interval)", "S", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 33) };
                    yield return new object[] { "+1 minute", "n", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 41, 34) };
                    yield return new object[] { "+0 minutes", "n", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 minute", "n", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 39, 34) };
                    yield return new object[] { "+1 hour", "h", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 22, 40, 34) };
                    yield return new object[] { "+0 hours", "h", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 hour", "h", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 20, 40, 34) };
                    yield return new object[] { "+1 weekday", "w", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 29, 21, 40, 34) };
                    yield return new object[] { "+0 weekday", "w", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 weekday", "w", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 27, 21, 40, 34) };
                    yield return new object[] { "+1 day", "d", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 29, 21, 40, 34) };
                    yield return new object[] { "+0 day", "d", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 day", "d", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 27, 21, 40, 34) };
                    yield return new object[] { "+1 day of year", "y", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 29, 21, 40, 34) };
                    yield return new object[] { "+0 day of year", "y", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 day of year", "y", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 27, 21, 40, 34) };
                    yield return new object[] { "+1 week", "ww", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 6, 4, 21, 40, 34) };
                    yield return new object[] { "+0 weeks", "ww", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 week", "ww", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 21, 21, 40, 34) };
                    yield return new object[] { "+1 month", "m", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 6, 28, 21, 40, 34) };
                    yield return new object[] { "+0 months", "m", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 month", "m", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 4, 28, 21, 40, 34) };
                    yield return new object[] { "+1 quarter", "q", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 8, 28, 21, 40, 34) };
                    yield return new object[] { "+0 quarters", "q", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 quarter", "q", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 2, 28, 21, 40, 34) };
                    yield return new object[] { "+1 year", "yyyy", 1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2016, 5, 28, 21, 40, 34) };
                    yield return new object[] { "+0 years", "yyyy", 0, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "-1 year", "yyyy", -1, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2014, 5, 28, 21, 40, 34) };

                    // Cases that illustrate how different numbers of dates in a month are dealt with
                    yield return new object[] { "+1 month from 31th Jan 2015 (will have to change date to 28th)", "m", 1, new DateTime(2015, 1, 31, 21, 40, 34), new DateTime(2015, 2, 28, 21, 40, 34) };
                    yield return new object[] { "+1 year from 29th Feb 2016 (will have to change date to 28th)", "yyyy", 1, new DateTime(2016, 2, 29, 21, 40, 34), new DateTime(2017, 2, 28, 21, 40, 34) };

                    // Unlike a lot of other VBScript functions, the arguments are only validated as they are required (in reverse order, which is consistent with other functions). So
                    // if the "value" argument is Null then it doesn't matter if the "number" argument is Nothing or the "interval" argument a non-supported value.
                    yield return new object[] { "Nothing \"number\" ignored if \"value\" is Null", "s", VBScriptConstants.Nothing, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Nothing \"interval\" ignored if \"value\" is Null", VBScriptConstants.Nothing, 1, DBNull.Value, DBNull.Value };
                    yield return new object[] { "Invalid \"interval\" ignored if \"value\" is Null", "x", 1, DBNull.Value, DBNull.Value };

                    // The fractional component is ignored from the "number" argument (unless it rolls over due to precision limitations - see PrecisionEdgeCases), it does NOT round up
                    yield return new object[] { "102.5 is treated as 102", "s", 102, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 42, 16) };
                    yield return new object[] { "102.9 is treated as 102", "s", 102.9, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 42, 16) };
                    yield return new object[] { "102.99 is treated as 102", "s", 102.99, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 42, 16) };
                    yield return new object[] { "103.5 is treated as 103", "s", 103, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 42, 17) };
                    yield return new object[] { "103.9 is treated as 103", "s", 103.9, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 42, 17) };
                    yield return new object[] { "103.99 is treated as 103", "s", 103.99, new DateTime(2015, 5, 28, 21, 40, 34), new DateTime(2015, 5, 28, 21, 42, 17) };

                    // VBScript's DateAdd has some crazy behaviour where any number outside the Int32 (VBScript's "Long") range gets treated as Int32.MinValue - ANY value. Bizarre.
                    yield return new object[] { "Int32.MaxValue acts as expected", "s", Int32.MaxValue, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MaxValue) };
                    yield return new object[] { "(Int32.MaxValue + 1) rolls over to Int32.MinValue", "s", (Int64)Int32.MaxValue + 1, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                    yield return new object[] { "(Int32.MaxValue + 2) is stuck on Int32.MinValue", "s", (Int64)Int32.MaxValue + 2, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                    yield return new object[] { "UInt32.MaxValue is stuck on Int32.MinValue", "s", UInt32.MaxValue, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                    yield return new object[] { "(UInt64.MaxValue * 10) is stuck on Int32.MinValue", "s", (double)UInt64.MaxValue * 10, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                    yield return new object[] { "(Int32.MinValue - 1) is treated as Int32.MinValue", "s", (Int64)Int32.MinValue - 1, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                    yield return new object[] { "(Int32.MinValue - 2) is treated as Int32.MinValue", "s", (Int64)Int32.MinValue - 2, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                    yield return new object[] { "Int64.MinValue is treated as Int32.MinValue", "s", Int64.MinValue, VBScriptConstants.ZeroDate, VBScriptConstants.ZeroDate.AddSeconds(Int32.MinValue) };
                }
            }

            public static IEnumerable<object[]> PrecisionEdgeCaseData
            {
                get
                {
                    yield return new object[] { "Edge of 1.999.. precision (15x 9s, treated as 1)", "h", 1, 15, new DateTime(2015, 5, 28, 21, 3, 52), new DateTime(2015, 5, 28, 22, 3, 52) };
                    yield return new object[] { "Edge of 1.999.. precision (16x 9s, treated as 2)", "h", 1, 16, new DateTime(2015, 5, 28, 21, 3, 52), new DateTime(2015, 5, 28, 23, 3, 52) };
                    yield return new object[] { "Edge of 10.999.. precision (15x 9s, treated as 10)", "h", 10, 15, new DateTime(2015, 5, 28, 21, 3, 52), new DateTime(2015, 5, 29, 7, 3, 52) };
                    yield return new object[] { "Edge of 10.999.. precision (16x 9s, treated as 11)", "h", 10, 16, new DateTime(2015, 5, 28, 21, 3, 52), new DateTime(2015, 5, 29, 8, 3, 52) };
                    yield return new object[] { "Edge of 100.999.. precision (14x 9s, treated as 100)", "h", 100, 14, new DateTime(2015, 5, 28, 21, 3, 52), new DateTime(2015, 6, 2, 1, 3, 52) };
                    yield return new object[] { "Edge of 100.999.. precision (15x 9s, treated as 101)", "h", 100, 15, new DateTime(2015, 5, 28, 21, 3, 52), new DateTime(2015, 6, 2, 2, 3, 52) };
                }
            }

            public static IEnumerable<object[]> TypeMismatchData
            {
                get
                {
                    yield return new object[] { "Non-numeric-or-null \"interval\"", "s", "x", new DateTime(2015, 5, 28, 21, 40, 34) };
               }
            }

            public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
            {
                get
                {
                    yield return new object[] { "Whitespace in \"interval\" is invalid", " s", 1, new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "Invalid interval blows up", "x", 1, new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "Date greater then allowable range", "d", 1, VBScriptConstants.LatestPossibleDate };
                    yield return new object[] { "Date before then allowable range", "d", -1, VBScriptConstants.EarliestPossibleDate };
               }
            }

            public static IEnumerable<object[]> InvalidUseOfNullData
            {
                get
                {
                    yield return new object[] { "Null \"number\" is invalid (unless \"value\" is Null)", "s", DBNull.Value, new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "Null \"interval\" is invalid (unless \"value\" is Null)", DBNull.Value, 1, new DateTime(2015, 5, 28, 21, 40, 34) };
                }
            }

            public static IEnumerable<object[]> ObjectVariableNotSetData
            {
                get
                {
                    yield return new object[] { "Nothing \"value\" is invalid", "s", 1, VBScriptConstants.Nothing };
                    yield return new object[] { "Nothing \"number\" is invalid (unless \"value\" is Null)", "s", VBScriptConstants.Nothing, new DateTime(2015, 5, 28, 21, 40, 34) };
                    yield return new object[] { "Nothing \"interval\" is invalid (unless \"value\" is Null)", VBScriptConstants.Nothing, 1, new DateTime(2015, 5, 28, 21, 40, 34) };
                }
            }
        }
    }
}
