using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		// TODO: Incomplete..
		public class DATEDIFF
		{
			[Theory, MemberData("SuccessData")]
			public void SuccessCases(string description, object interval, object date1, object date2, object expectedResult)
			{
				Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().DATEDIFF(interval, date1, date2));
			}

			[Theory, MemberData("TypeMismatchData")]
			public void TypeMismatchCases(string description, object interval, object date1, object date2)
			{
				Assert.Throws<TypeMismatchException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DATEDIFF(interval, date1, date2);
				});
			}

			[Theory, MemberData("InvalidProcedureCallOrArgumentData")]
			public void InvalidProcedureCallOrArgumentCases(string description, object interval, object date1, object date2)
			{
				Assert.Throws<InvalidProcedureCallOrArgumentException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DATEDIFF(interval, date1, date2);
				});
			}

			[Theory, MemberData("InvalidUseOfNullData")]
			public void InvalidUseOfNullCases(string description, object interval, object date1, object date2)
			{
				Assert.Throws<InvalidUseOfNullException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DATEDIFF(interval, date1, date2);
				});
			}

			[Theory, MemberData("ObjectVariableNotSetData")]
			public void ObjectVariableNotSetCases(string description, object interval, object date1, object date2)
			{
				Assert.Throws<ObjectVariableNotSetException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().DATEDIFF(interval, date1, date2);
				});
			}

			public static IEnumerable<object[]> SuccessData
			{
				get
				{
					// Note: We could go to town with test cases for the various string formats that VBScript supports, but the DATEDIFFimplementation backs onto the DateParser and
					// it would be duplication of effort going through everything again here

					// Middle-of-the-road cases
					// TODO: This is only a very small subset (that I've added to cover the cases that I need immediately in my translated code - but the test suite should be fully in the future)
					yield return new object[] { "+1 day", "d", new DateTime(2017, 2, 22), new DateTime(2017, 2, 23), 1 };
					yield return new object[] { "+1 day (from 23:59:59 to 00:00:00 next day)", "d", new DateTime(2017, 2, 22, 23, 59, 59), new DateTime(2017, 2, 23, 0, 0, 0), 1 };
					yield return new object[] { "-1 day", "d", new DateTime(2017, 2, 22), new DateTime(2017, 2, 21), -1 };

					yield return new object[] { "-1 day --TODO1", "m", new DateTime(2017, 1, 1), new DateTime(2017, 1, 16), 0 };
					yield return new object[] { "-1 day --TODO2", "m", new DateTime(2017, 2, 1), new DateTime(2017, 1, 16), -1 };
					yield return new object[] { "-1 day --TODO3", "m", new DateTime(2017, 1, 1), new DateTime(2017, 2, 16), 1 };
					yield return new object[] { "-1 day --TODO4", "m", new DateTime(2017, 1, 1), new DateTime(2016, 2, 16), -11 };
				}
			}

			public static IEnumerable<object[]> TypeMismatchData
			{
				get
				{
					return new object[0][]; // TODO
				}
			}

			public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
			{
				get
				{
					return new object[0][]; // TODO
				}
			}

			public static IEnumerable<object[]> InvalidUseOfNullData
			{
				get
				{
					return new object[0][]; // TODO
				}
			}

			public static IEnumerable<object[]> ObjectVariableNotSetData
			{
				get
				{
					return new object[0][]; // TODO
				}
			}
		}
	}
}
