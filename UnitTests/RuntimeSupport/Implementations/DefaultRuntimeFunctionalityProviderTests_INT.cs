using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		public class INT
		{
			[Theory, MemberData("SuccessData")]
			public void SuccessCases(string description, object value, object expectedResult)
			{
				Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().INT(value));
			}

			[Theory, MemberData("ObjectVariableNotSetData")]
			public void ObjectVariableNotSetCases(string description, object value)
			{
				Assert.Throws<ObjectVariableNotSetException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().INT(value);
				});
			}

			[Theory, MemberData("TypeMismatchData")]
			public void TypeMismatchCases(string description, object value)
			{
				Assert.Throws<TypeMismatchException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().INT(value);
				});
			}

			[Theory, MemberData("ObjectDoesNotSupportPropertyOrMemberData")]
			public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object value)
			{
				Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().INT(value);
				});
			}

			public static IEnumerable<object[]> SuccessData
			{
				get
				{
					return new[]
					{
						new object[] { "Empty", null, (Int16)0 },
						new object[] { "Null", DBNull.Value, DBNull.Value },
						new object[] { "True", true, (Int16)(-1) },
						new object[] { "False", false, (Int16)0 },
						new object[] { "Integer", (Int16)123, (Int16)123 },
						new object[] { "Long (within Integer range)", (Int32)123, (Int32)123 },
						new object[] { "Single (within Integer range)", (Single)123, (Single)123 },
						new object[] { "Double (within Integer range)", (Double)123, (Double)123 },
						new object[] { "Decimal (within Integer range)", (Decimal)123, (Decimal)123 },
						new object[] { "Date (removes time component)", new DateTime(2017, 3, 8, 18, 30, 12, 22), new DateTime(2017, 3, 8) },
						new object[] { "String representing numeric value", "123.45", (double)123 },
						new object[] { "String representing numeric value with leading and trailing whitespace", " 123.45 ", (double)123 },
						new object[] { "Object with default property which is decimal 123.45", new exampledefaultpropertytype { result = 123.45m }, 123m },

						// A few tests to reinforce that the fraction is removed, it's NOT rounded away from zero or even numbers
						new object[] { "0.5", (double)0.5, (double)0 },
						new object[] { "1.5", (double)1.5, (double)1 },
						new object[] { "2.5", (double)2.5, (double)2 },
						new object[] { "3.5", (double)3.5, (double)3 },

						// These results are surprising, I had expected VBScript to remove the fraction from a number like -0.5 to leave 0 (or from -1.5 to leave -1) but it doesn't!
						new object[] { "-0.5", (double)(-0.5), (double)(-1) },
						new object[] { "-1.5", (double)(-1.5), (double)(-2) },
						new object[] { "-2.5", (double)(-2.5), (double)(-3) },
						new object[] { "-3.5", (double)(-3.5), (double)(-4) }
					};
				}
			}

			public static IEnumerable<object[]> ObjectVariableNotSetData
			{
				get
				{
					return new[] { new object[] { "Nothing", VBScriptConstants.Nothing } };
				}
			}

			public static IEnumerable<object[]> TypeMismatchData
			{
				get
				{
					return new[]
					{
						new object[] { "Blank String", "" },
						new object[] { "Whitespace", " " },
						new object[] { "String representation of boolean", "True" },
						new object[] { "String representing of numeric value whitespace around decimal point", "123. 45" },
						new object[] { "Unintialised array", new object[0] }
					};
				}
			}

			public static IEnumerable<object[]> ObjectDoesNotSupportPropertyOrMemberData
			{
				get
				{
					return new[] { new  object[] { "Object without default property", new object() } };
				}
			}
		}
	}
}
