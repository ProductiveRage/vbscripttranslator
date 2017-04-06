using System;
using System.Collections.Generic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	public static partial class DefaultRuntimeFunctionalityProviderTests
	{
		public class HEX
		{
			[Theory, MemberData("SuccessData")]
			public void SuccessCases(string description, object value, object expectedResult)
			{
				Assert.Equal(expectedResult, DefaultRuntimeSupportClassFactory.Get().HEX(value));
			}

			[Theory, MemberData("ObjectVariableNotSetData")]
			public void ObjectVariableNotSetCases(string description, object value)
			{
				Assert.Throws<ObjectVariableNotSetException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().HEX(value);
				});
			}

			[Theory, MemberData("TypeMismatchData")]
			public void TypeMismatchCases(string description, object value)
			{
				Assert.Throws<TypeMismatchException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().HEX(value);
				});
			}

			[Theory, MemberData("ObjectDoesNotSupportPropertyOrMemberData")]
			public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object value)
			{
				Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().HEX(value);
				});
			}

			[Theory, MemberData("OverflowData")]
			public void OverflowCases(string description, object value)
			{
				Assert.Throws<VBScriptOverflowException>(() =>
				{
					DefaultRuntimeSupportClassFactory.Get().HEX(value);
				});
			}

			public static IEnumerable<object[]> SuccessData
			{
				get
				{
					return new[]
					{
						// Unlike some functions, Null IS acceptable
						new object[] { "Null", DBNull.Value, DBNull.Value },

						// Zero-like values
						new object[] { "Empty", null, "0" },
						new object[] { "0 (Integer)", (short)0, "0" },
						new object[] { "0 (Double)", 0d, "0" },
						new object[] { "False", false, "0" },

						// Larger positive values
						new object[] { "1 (Byte)", (byte)1, "1" },
						new object[] { "1 (Integer)", (short)1, "1" },
						new object[] { "1 (Currency)", 1m, "1" },
						new object[] { "1 (Single)", 1f, "1" },
						new object[] { "32767 (Integer)", (short)32767, "7FFF" },
						new object[] { "32768 (Long)", 32768, "8000" },
						new object[] { "2147483647 (Long)", 2147483647, "7FFFFFFF" }, // Largest positive numer acceptable before overflow

						// -1 values
						new object[] { "-1 (Integer)", (short)(-1), "FFFF" },
						new object[] { "-2 (Integer)", (short)(-2), "FFFE" },
						new object[] { "-1 (Double)", -1d, "FFFFFFFF" },
						new object[] { "-2 (Double)", -2d, "FFFFFFFE" },
						new object[] { "-1 (String)", "-1", "FFFFFFFF" },
						new object[] { "True", true, "FFFF" },

						// Larger negative values
						new object[] { "-32767 (Integer)", (short)(-32767), "8001" },
						new object[] { "-32768 (Long)", -32768, "FFFF8000" },
						new object[] { "-2147483648 (Double)", -2147483648d, "80000000" }, // Largest negative numer acceptable before overflow

						// A few tests to reinforce that the rounding of numbers works as required
						new object[] { "0.1 (Double)", 0.1d, "0" },
						new object[] { "0.4 (Double)", 0.4d, "0" },
						new object[] { "0.5 (Double)", 0.5d, "0" },
						new object[] { "0.6 (Double)", 0.6d, "1" },
						new object[] { "1.1 (Double)", 1.1d, "1" },
						new object[] { "1.4 (Double)", 1.4d, "1" },
						new object[] { "1.5 (Double)", 1.5d, "2" },
						new object[] { "1.6 (Double)", 1.6d, "2" },
						new object[] { "2.1 (Double)", 2.1d, "2" },
						new object[] { "2.4 (Double)", 2.4d, "2" },
						new object[] { "2.5 (Double)", 2.5d, "2" },
						new object[] { "2.6 (Double)", 2.6d, "3" },
						new object[] { "3.1 (Double)", 3.1d, "3" },
						new object[] { "3.4 (Double)", 3.4d, "3" },
						new object[] { "3.5 (Double)", 3.5d, "4" },
						new object[] { "3.6 (Double)", 3.6d, "4" },
						new object[] { "-0.1 (Double)", -0.1d, "0" },
						new object[] { "-0.4 (Double)", -0.4d, "0" },
						new object[] { "-0.5 (Double)", -0.5d, "0" },
						new object[] { "-0.6 (Double)", -0.6d, "FFFFFFFF" },
						new object[] { "-1 (Double)", -1d, "FFFFFFFF" },
						new object[] { "-1.1 (Double)", -1.1d, "FFFFFFFF" },
						new object[] { "-1.4 (Double)", -1.4d, "FFFFFFFF" },
						new object[] { "-1.5 (Double)", -1.5d, "FFFFFFFE" },
						new object[] { "-1.6 (Double)", -1.6d, "FFFFFFFE" },
						new object[] { "-2.1 (Double)", -2.1d, "FFFFFFFE" },
						new object[] { "-2.4 (Double)", -2.4d, "FFFFFFFE" },
						new object[] { "-2.5 (Double)", -2.5d, "FFFFFFFE" },
						new object[] { "-2.6 (Double)", -2.6d, "FFFFFFFD" },
						new object[] { "-3.1 (Double)", -3.1d, "FFFFFFFD" },
						new object[] { "-3.4 (Double)", -3.4d, "FFFFFFFD" },
						new object[] { "-3.5 (Double)", -3.5d, "FFFFFFFC" },
						new object[] { "-3.6 (Double)", -3.6d, "FFFFFFFC" }
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

			public static IEnumerable<object[]> OverflowData
			{
				get
				{
					return new[]
					{
						new object[] { "2147483648", 2147483648 },
						new object[] { "-2147483649", -2147483649 }
					};
				}
			}
		}
	}
}
