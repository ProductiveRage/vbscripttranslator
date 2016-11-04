using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Attributes;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		/// <summary>
		/// This is EXTREMELY close to CDATE.. but not exactly the same (the Null case is the only difference, I think - that CDATE throws an error while this returns Null)
		/// </summary>
		public class WEEKDAY
		{
			public WEEKDAY()
			{
				// WEEKDAY uses CurrentCulture to figure out the first day of the week when vbUseSystemDayOfWeek is passed to it
				// An "en-GB" culture is set here so that our tests can assume Monday is the first day of the week when testing vbUseSystemDayOfWeek
				Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-GB");
				Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo("en-GB");
			}

			[Theory, MemberData("SuccessData")]
			public void SuccessCases(string description, object value, object firstDayOfWeek, object expectedResult)
			{
				Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().WEEKDAY(value, firstDayOfWeek));
			}
			
			[Theory, MemberData("TypeMismatchData")]
			public void TypeMismatchCases(string description, object value, object firstDayOfWeek)
			{
				Assert.Throws<TypeMismatchException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().WEEKDAY(value, firstDayOfWeek);
				});
			}

			[Theory, MemberData("ObjectVariableNotSetData")]
			public void ObjectVariableNotSetCases(string description, object value, object firstDayOfWeek)
			{
				Assert.Throws<ObjectVariableNotSetException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().WEEKDAY(value, firstDayOfWeek);
				});
			}

			[Theory, MemberData("OverflowData")]
			public void OverflowCases(string description, object value, object firstDayOfWeek)
			{
				Assert.Throws<VBScriptOverflowException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().WEEKDAY(value, firstDayOfWeek);
				});
			}

			[Theory, MemberData("InvalidProcedureCallOrArgumentData")]
			public void InvalidProcedureCallOrArgumentCases(string description, object value, object firstDayOfWeek)
			{
				Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().WEEKDAY(value, firstDayOfWeek);
				});
			}

			/// <summary>
			/// The default firstDayOfWeek argument to Weekday if not specified is vbSunday.
			/// </summary>
			private static readonly int DefaultFirstDayOfWeek = VBScriptConstants.vbSunday;

			public static IEnumerable<object[]> SuccessData
			{
				get
				{
					// Note: Empty date = 1899-12-30 (a Saturday)
					yield return new object[] { "Empty",                          null,                                 DefaultFirstDayOfWeek, VBScriptConstants.vbSaturday };
					yield return new object[] { "Null",                           DBNull.Value,                         DefaultFirstDayOfWeek, DBNull.Value };
					yield return new object[] { "Zero",                           0,                                    DefaultFirstDayOfWeek, VBScriptConstants.vbSaturday };
					yield return new object[] { "Minus one",                      -1,                                   DefaultFirstDayOfWeek, VBScriptConstants.vbFriday };
					yield return new object[] { "Minus 400",                      -400,                                 DefaultFirstDayOfWeek, VBScriptConstants.vbFriday };
					yield return new object[] { "Plus 40000",                     40000,                                DefaultFirstDayOfWeek, VBScriptConstants.vbMonday };
					yield return new object[] { "String \"-400.2\"",              "-400.2",                             DefaultFirstDayOfWeek, VBScriptConstants.vbFriday };
					yield return new object[] { "String \"40000.2\"",             "40000.2",                            DefaultFirstDayOfWeek, VBScriptConstants.vbMonday };
					yield return new object[] { "String \"2009-10-11\"",          "2009-10-11",                         DefaultFirstDayOfWeek, VBScriptConstants.vbSunday };
					yield return new object[] { "String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44",                DefaultFirstDayOfWeek, VBScriptConstants.vbSunday };
					yield return new object[] { "A Date",                         new DateTime(2009, 7, 6, 20, 12, 44), DefaultFirstDayOfWeek, VBScriptConstants.vbMonday };

					// Note: When overriding the firstDayOfWeek, I'm using literal numbers directly to represent the return values because
					// the enum names don't make sense in that case (cause the values are shifted)
					// Note: The tests will be run with en-GB CurrentCulture, so vbUseSystemDayOfWeek will use Monday as the first day of the week
					yield return new object[] { "A Date with Monday week start",    new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbMonday,             1 };
					yield return new object[] { "A Date with Tuesday week start",   new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbTuesday,            7 };
					yield return new object[] { "A Date with Wednesday week start", new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbWednesday,          6 };
					yield return new object[] { "A Date with Thursday week start",  new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbThursday,           5 };
					yield return new object[] { "A Date with Friday week start",    new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbFriday,             4 };
					yield return new object[] { "A Date with Saturday week start",  new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbSaturday,           3 };
					yield return new object[] { "A Date with Sunday week start",    new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbSunday,             2 };
					yield return new object[] { "A Date with *System* week start",  new DateTime(2009, 7, 6, 20, 12, 44), VBScriptConstants.vbUseSystemDayOfWeek, 1 };

					yield return new object[] { "Object with default property which is Empty",                          new exampledefaultpropertytype(),                                  DefaultFirstDayOfWeek, VBScriptConstants.vbSaturday };
					yield return new object[] { "Object with default property which is Null",                           new exampledefaultpropertytype { result = DBNull.Value },          DefaultFirstDayOfWeek, DBNull.Value };
					yield return new object[] { "Object with default property which is Zero",                           new exampledefaultpropertytype { result = 0 },                     DefaultFirstDayOfWeek, VBScriptConstants.vbSaturday };
					yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, DefaultFirstDayOfWeek, VBScriptConstants.vbSunday };

					// Overflow edge checks
					yield return new object[] { "Largest positive integer before overflow", 2958465, DefaultFirstDayOfWeek, VBScriptConstants.vbFriday };
					yield return new object[] { "Largest negative integer before overflow", -657434, DefaultFirstDayOfWeek, VBScriptConstants.vbFriday };
				}
			}

			public static IEnumerable<object[]> TypeMismatchData
			{
				get
				{
					yield return new object[] { "1st arg - Blank string", "", DefaultFirstDayOfWeek };
					yield return new object[] { "1st arg - Object with default property which is a blank string", new exampledefaultpropertytype { result = "" }, DefaultFirstDayOfWeek };
					yield return new object[] { "2nd arg - Blank string", null, "" };
					yield return new object[] { "2nd arg - Object with default property which is a blank string", null, new exampledefaultpropertytype { result = "" } };
				}
			}

			public static IEnumerable<object[]> ObjectVariableNotSetData
			{
				get
				{
					yield return new object[] { "1st arg - Nothing", VBScriptConstants.Nothing, DefaultFirstDayOfWeek };
					yield return new object[] { "1st arg - Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing }, DefaultFirstDayOfWeek };
					yield return new object[] { "2nd arg - Nothing", null, VBScriptConstants.Nothing };
					yield return new object[] { "2nd arg - Object with default property which is Nothing", null, new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
				}
			}

			public static IEnumerable<object[]> OverflowData
			{
				get
				{
					yield return new object[] { "1st arg - Large number (12388888888888.2)", 12388888888888.2, DefaultFirstDayOfWeek };
					yield return new object[] { "1st arg - Object with default property which is a large number (12388888888888.2)", new exampledefaultpropertytype { result = 12388888888888.2 }, DefaultFirstDayOfWeek };
					yield return new object[] { "1st arg - Smallest positive integer that overflows", 2958466, DefaultFirstDayOfWeek };
					yield return new object[] { "1st arg - Smallest negative integer that overflows", -657435, DefaultFirstDayOfWeek };
					yield return new object[] { "2nd arg - Smallest positive integer that overflows", 2147483648, DefaultFirstDayOfWeek };
					yield return new object[] { "2nd arg - Smallest negative integer that overflows", -2147483649, DefaultFirstDayOfWeek };
				}
			}

			public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
			{
				get
				{
					yield return new object[] { "2nd arg - -1", null, -1 };
					yield return new object[] { "2nd arg - Object with default property which is -1", null, new exampledefaultpropertytype { result = -1 } };
					yield return new object[] { "2nd arg - 8", null, 8 };
					yield return new object[] { "2nd arg - Object with default property which is 8", null, new exampledefaultpropertytype { result = 8 } };
				}
			}
		}
	}
}
