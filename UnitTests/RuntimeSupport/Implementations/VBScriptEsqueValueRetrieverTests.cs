using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Attributes;
using VBScriptTranslator.RuntimeSupport.Compat;
using VBScriptTranslator.RuntimeSupport.Exceptions;
using VBScriptTranslator.RuntimeSupport.Implementations;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
	// TODO: Add a test that complements ExpressionGeneratorTests.PropertyAccessOnStringLiteralResultsInRuntimeError, that confirms that if a string value is
	// provided as call target that an error will be raised for property or method access attempts
	public class VBScriptEsqueValueRetrieverTests
	{
		[Fact]
		public void VALDoesNotAlterNull()
		{
			Assert.Equal(
				null,
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(null)
			);
		}

		[Fact]
		public void VALDoesNotAlterOne()
		{
			Assert.Equal(
				1,
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1)
			);
		}

		[Fact]
		public void VALDoesNotAlterOnePointOneFloat()
		{
			Assert.Equal(
				1.1f,
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1.1f)
			);
		}

		[Fact]
		public void VALDoesNotAlterOnePointOneDouble()
		{
			Assert.Equal(
				1.1d,
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1.1d)
			);
		}

		[Fact]
		public void VALDoesNotAlterOnePointOneDecimal()
		{
			Assert.Equal(
				1.1m,
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1.1m)
			);
		}

		[Fact]
		public void VALDoesNotAlterMinusOne()
		{
			Assert.Equal(
				-1,
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(-1)
			);
		}

		[Fact]
		public void VALDoesNotAlterEmptyString()
		{
			Assert.Equal(
				"",
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL("")
			);
		}

		[Fact]
		public void VALDoesNotAlterNonEmptyString()
		{
			Assert.Equal(
				"Test",
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL("Test")
			);
		}

		[Fact]
		public void VALFailsOnTranslatedClassWithNoDefaultMember()
		{
			// Execute twice to ensure that the TryVAL caching does not affect the result
			Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new translatedclasswithnodefaultmember()));
			Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new translatedclasswithnodefaultmember()));
		}

		[SourceClassName("TranslatedClassWithNoDefaultMember")]
		private class translatedclasswithnodefaultmember { }

		[Fact]
		public void VALFailsOnComObjectWithNoParameterlessDefaultMember()
		{
			// Execute twice to ensure that the TryVAL caching does not affect the result
			var dictionary = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
			Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(dictionary));
			Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(dictionary));
		}

		/// <summary>
		/// This test relates to a fix just applied to GenerateSetInvoker (where an argumentsArray reference was being used instead of invokeArguments, which meant that the same set of arguments
		/// were being reused on each call)
		/// </summary>
		[Fact]
		public void EnsureThatOldArgumentsAreNotReusedInSubsequentIDispatchCalls()
		{
			// This requires that the project be built in 32-bit mode (as much of the IDispatch support does)
			var dict = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
			using (var _ = VBScriptTranslator.RuntimeSupport.DefaultRuntimeSupportClassFactory.Get())
			{
				_.SET(1, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("a"));
				_.SET(2, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("b"));
				Assert.Equal(2, _.CALL(context: null, target: dict, member1: "Count"));
			}
		}

		[Fact]
		public void VALFailsOnNonComVisibleNonTranslatedClasses()
		{
			// Execute twice to ensure that the TryVAL caching does not affect the result
			Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new NonComVisibleNonTranslatedClass()));
			Assert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() => DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new NonComVisibleNonTranslatedClass()));
		}

		private class NonComVisibleNonTranslatedClass { }

		[Fact]
		public void VALSupportsIsDefaultAttributeOnTranslatedClasses()
		{
			// Execute twice to ensure that the TryVAL caching does not affect the result
			Assert.Equal("name!", DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new translatedclasswithdefaultmember()));
			Assert.Equal("name!", DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new translatedclasswithdefaultmember()));
		}

		[SourceClassName("TranslatedClassWithNoDefaultMember")]
		private class translatedclasswithdefaultmember
		{
			[IsDefault]
			public string name() { return "name!"; }
		}

		[Fact]
		public void VALSupportsDefaultMemberAttributeOnComVisibleNonTranslatedClasses()
		{
			// Execute twice to ensure that the TryVAL caching does not affect the result
			Assert.Equal("name!", DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithDefaultMember()));
			Assert.Equal("name!", DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithDefaultMember()));
		}

		[ComVisible(true)]
		[DefaultMember("Name")]
		private class ComVisibleNonTranslatedClassWithDefaultMember
		{
			public string Name { get { return "name!"; } }
		}

		[Fact]
		public void VALSupportsToStringOnComVisibleNonTranslatedClasses()
		{
			// Execute twice to ensure that the TryVAL caching does not affect the result
			var target = new ComVisibleNonTranslatedClassWithDefaultMember();
			Assert.Equal(
				"VBScriptTranslator.UnitTests.RuntimeSupport.Implementations.VBScriptEsqueValueRetrieverTests+ComVisibleNonTranslatedClassWithNoDefaultMember",
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithNoDefaultMember())
			);
			Assert.Equal(
				"VBScriptTranslator.UnitTests.RuntimeSupport.Implementations.VBScriptEsqueValueRetrieverTests+ComVisibleNonTranslatedClassWithNoDefaultMember",
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(new ComVisibleNonTranslatedClassWithNoDefaultMember())
			);
		}

		[ComVisible(true)]
		private class ComVisibleNonTranslatedClassWithNoDefaultMember { }

		[Fact]
		public void IFOfNullIsFalse()
		{
			Assert.False(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(null)
			);
		}

		[Fact]
		public void IFOfZeroIsFalse()
		{
			Assert.False(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(0)
			);
		}

		[Fact]
		public void IFOfOneIsTrue()
		{
			Assert.True(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(1)
			);
		}

		[Fact]
		public void IFOfMinusOneIsTrue()
		{
			Assert.True(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(-1)
			);
		}

		[Fact]
		public void IFOfOnePointOneIsTrue()
		{
			Assert.True(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(1.1)
			);
		}

		/// <summary>
		/// VBScript doesn't round the number down to zero and find 0.1 to be false, it just checks that the number is non-zero
		/// </summary>
		[Fact]
		public void IFOfPointOneIsTrue()
		{
			Assert.True(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(0.1)
			);
		}

		[Fact]
		public void IFOfPointStringRepresentationOfOneIsTrue()
		{
			Assert.True(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("1")
			);
		}

		[Fact]
		public void IFThrowsExceptionForBlanksString()
		{
			Assert.Throws<TypeMismatchException>(() =>
			{
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("");
			});
		}

		[Fact]
		public void IFThrowsExceptionForNonNumericString()
		{
			Assert.Throws<TypeMismatchException>(() =>
			{
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("one");
			});
		}

		[Fact]
		public void IFIgnoresWhiteSpaceWhenParsingStrings()
		{
			Assert.True(
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("   1    ")
			);
		}

		[Fact]
		public void NativeClassSupportsMethodCallWithArguments()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				new PseudoField { value = "value:F1" },
				_.CALL(
					context: null,
					target: new PseudoRecordset(),
					members: new[] { "fields" },
					argumentProviderBuilder: _.ARGS.Val("F1")
				),
				new PseudoFieldObjectComparer()
			);
		}

		[Fact]
		public void NativeClassSupportsDefaultMethodCallWithArguments()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				new PseudoField { value = "value:F1" },
				_.CALL(
					context: null,
					target: new PseudoRecordset(),
					members: new string[0],
					argumentProviderBuilder: _.ARGS.Val("F1")
				),
				new PseudoFieldObjectComparer()
			);
		}

		[Fact]
		public void ADORecordsetSupportsNamedFieldAccess()
		{
			var recordset = new ADODB.Recordset();
			recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
			recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
			recordset.AddNew();
			recordset.Fields["name"].Value = "TestName";
			recordset.Update();

			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				recordset.Fields["name"],
				_.CALL(
					context: null,
					target: recordset,
					members: new[] { "fields" },
					argumentProviderBuilder: _.ARGS.Val("name")
				),
				new ADOFieldObjectComparer()
			);
		}

		[Fact]
		public void ADORecordsetSupportsDefaultFieldAccess()
		{
			var recordset = new ADODB.Recordset();
			recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
			recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
			recordset.AddNew();
			recordset.Fields["name"].Value = "TestName";
			recordset.Update();

			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				recordset.Fields["name"],
				_.CALL(
					context: null,
					target: recordset,
					members: new string[0],
					argumentProviderBuilder: _.ARGS.Val("name")
				),
				new ADOFieldObjectComparer()
			);
		}

		/// <summary>
		/// This describes an extremely common VBScript pattern - rstResults("name") needs to return a value by using the default Fields access
		/// (passing through "name" to it) and then default access of the Field's Value property (which requires a VAL call, which must be
		/// included in translated code if a value type is expected - any time other than when a SET statement is present as part of a
		/// variable assignment)
		/// </summary>
		[Fact]
		public void ADORecordsetSupportsDefaultFieldValueAccess()
		{
			var recordset = new ADODB.Recordset();
			recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
			recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
			recordset.AddNew();
			recordset.Fields["name"].Value = "TestName";
			recordset.Update();

			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				"TestName",
				_.VAL(
					_.CALL(
						context: null,
						target: recordset,
						members: new string[0],
						argumentProviderBuilder: _.ARGS.Val("name")
					)
				)
			);
		}

		[Fact]
		public void OneDimensionalArrayAccessIsSupported()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			var data = new object[] { "One" };
			Assert.Equal(
				"One",
				_.CALL(
					context: null,
					target: data,
					members: new string[0],
					argumentProviderBuilder: _.ARGS.Val("0")
				)
			);
		}

		[Fact]
		public void ByRefArgumentIsUpdatedAfterCall()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			object arg0 = 1;
			_.CALL(context: null, target: this, member1: "ByRefArgUpdatingFunction", argumentProviderBuilder: _.ARGS.Ref(arg0, v => { arg0 = v; }).Val(false));
			Assert.Equal(123, arg0);
		}

		[Fact]
		public void ByRefArgumentIsUpdatedAfterCallEvenIfExceptionIsThrown()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			object arg0 = 1;
			try
			{
				_.CALL(context: null, target: this, member1: "ByRefArgUpdatingFunction", argumentProviderBuilder: _.ARGS.Ref(arg0, v => { arg0 = v; }).Val(true));
			}
			catch { }
			Assert.Equal(123, arg0);
		}

		[Fact]
		public void SingleArgumentParamsArrayMethodMayBeCalledWithZeroValues()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(0, _.CALL(context: null, target: this, member1: "GetNumberOfArgumentsPassedInParamsObjectArray", argumentProviderBuilder: _.ARGS));
		}

		[Fact]
		public void SingleArgumentParamsArrayMethodMayBeCalledWithSingleValue()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(1, _.CALL(context: null, target: this, member1: "GetNumberOfArgumentsPassedInParamsObjectArray", argumentProviderBuilder: _.ARGS.Val(1)));
		}

		[Fact]
		public void SingleArgumentParamsArrayMethodMayBeCalledWithTwoValues()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(2, _.CALL(context: null, target: this, member1: "GetNumberOfArgumentsPassedInParamsObjectArray", argumentProviderBuilder: _.ARGS.Val(1).Val(2)));
		}

		public void ByRefArgUpdatingFunction(ref object arg0, bool throwExceptionAfterUpdatingArgument)
		{
			arg0 = 123;
			if (throwExceptionAfterUpdatingArgument)
				throw new Exception("Example exception");
		}

		public Nullable<int> GetNumberOfArgumentsPassedInParamsObjectArray(params object[] args)
		{
			return (args == null) ? (int?)null : args.Length;
		}

		/// <summary>
		/// When a CALL execution is generated by the translator, the string member accessors should not be renamed - to avoid C# keywords, for example.
		/// If the target is a translated class then any members that use reserved C# keywords must be renamed, but if the target is not a translated
		/// class (a COM component, for example) then the call will fail, so the renaming must not be done at translation time. This means that the
		/// CALL implementation must support name rewriting for when the target is a translated-from-VBScript C# class.
		/// </summary>
		[Fact]
		public void StringMemberAccessorValuesShouldNotBeRewrittenAtTranslationTime()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				"Success!",
				_.CALL(context: null, target: new ImpressionOfTranslatedClassWithRewrittenPropertyName(), member1: "Params")
			);
		}

		[Fact]
		public void DelegateWithIncorrectNumberOfArguments()
		{
			var parameterLessDelegate = (Func<object>)(() => "delegate result");
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Throws<TargetParameterCountException>(
				() => _.CALL(context: null, target: parameterLessDelegate, members: new string[0], argumentProvider: _.ARGS.Val(1).GetArgs())
			);
		}

		/// <summary>
		/// In C#, it's fine to access an index within a string since a string is an array of characters. But in VBScript, it's not.
		/// </summary>
		[Fact]
		public void ItIsNotValidToAccessStringValueWithArguments()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Throws<TypeMismatchException>(
				() => _.CALL(context: null, target: "abc", members: new string[0], argumentProvider: _.ARGS.Val(0).GetArgs())
			);
		}

		/// <summary>
		/// DispId(0) is only supported when the match is unambiguous - previously, DispId(0) being specified on a property and on the getter for
		/// that property was considered an ambiguous match, but that shouldn't be the case since they both effectively refer to the same thing
		/// </summary>
		[Fact]
		public void SupportDispIdoZeroBeingRepeatedOnPropertyAndOnPropertyGetterWhenDefaultMemberRequired()
		{
			var target = new DispIdZeroRepeatedOnPropertyAndItsGetter("test");
			Assert.Equal(
				"test",
				DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(target)
			);
		}

		[Fact]
		public void DispIdZeroPropertySettingWorksWithValueTypes()
		{
			// This requires that the project be built in 32-bit mode (as much of the IDispatch support does)
			var dict = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
			var valueTypeValueToRecord = 123;
			using (var _ = VBScriptTranslator.RuntimeSupport.DefaultRuntimeSupportClassFactory.Get())
			{
				_.SET(valueTypeValueToRecord, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("ACCO"));
			}
		}

		[Fact]
		public void DispIdZeroPropertySettingWorksWithReferenceTypes()
		{
			// This requires that the project be built in 32-bit mode (as much of the IDispatch support does)
			var dict = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
			var referenceTypeValueToRecord = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
			using (var _ = VBScriptTranslator.RuntimeSupport.DefaultRuntimeSupportClassFactory.Get())
			{
				_.SET(referenceTypeValueToRecord, context: dict, target: dict, optionalMemberAccessor: null, argumentProviderBuilder: _.ARGS.Val("ACCO"));
			}
		}

		[ComVisible(true)]
		private class DispIdZeroRepeatedOnPropertyAndItsGetter
		{
			private readonly string _name;
			public DispIdZeroRepeatedOnPropertyAndItsGetter(string name)
			{
				_name = name;
			}

			[DispId(0)]
			public string Name
			{
				[DispId(0)]
				get { return _name; }
			}
		}


		/*====================================================================================================*/
		[Fact]
		public void GettingNestedObjectDefaultIndexedPropertyWorks()
		{
			// This tests that calls like `Session.Contents(...)` works. Contents is a regular getter-only property on Session.
			// The brackets in there are used because Contents' type is a dictionary that has a default property with an index parameter.
			// So essentially that code could also be written as `(Session.Contents)(...)` to make it clear that the Contents object
			// itself has the indexer, not the Contents property on Session. (Case 33958)
			var obj = new ExampleClassWithRegularPropertyThatHasADefaultIndexedPropertyWithin();
			string key = "ABC";
			string val = "DEF";
			obj.PropWithDefaultIndexedProp[key] = val;
			string result;
			using (var _ = DefaultRuntimeSupportClassFactory.Get())
			{
				result = _.VAL(_.CALL(context: null, target: obj, member1: "PropWithDefaultIndexedProp", argumentProviderBuilder: _.ARGS.Val(key))) as string;
			}
			Assert.True(result == "DEF");
		}

		[Fact]
		public void SettingNestedObjectDefaultIndexedPropertyWorks()
		{
			// This tests that calls like `Session.Contents(...) = ...` works. Contents is a regular getter-only property on Session.
			// The brackets in there are used because Contents' type is a dictionary that has a default property with an index parameter.
			// So essentially that code could also be written as `(Session.Contents)(...) = ...` to make it clear that the Contents object
			// itself has the indexer, not the Contents property on Session. (Case 33958)
			var obj = new ExampleClassWithRegularPropertyThatHasADefaultIndexedPropertyWithin();
			string newKey = "ABC";
			string newVal = "DEF";
			using (var _ = DefaultRuntimeSupportClassFactory.Get())
			{
				_.SET(newVal, context: null, target: obj, optionalMemberAccessor: "PropWithDefaultIndexedProp", argumentProviderBuilder: _.ARGS.Val(newKey));
			}
			Assert.True(obj.PropWithDefaultIndexedProp["ABC"] == "DEF");
		}

		[ComVisible(true)]
		private class ExampleClassWithRegularPropertyThatHasADefaultIndexedPropertyWithin
		{
			public ExampleClassWithDefaultIndexedProperty PropWithDefaultIndexedProp { get; } = new ExampleClassWithDefaultIndexedProperty();
		}

		[ComVisible(true)]
		private class ExampleClassWithDefaultIndexedProperty
		{
			private Dictionary<string, string> _dict = new Dictionary<string, string>();

			[DispId(0)]
			public string this[string key] {
				get { return _dict[key]; }
				set { _dict[key] = value; }
			}
		}
		/*====================================================================================================*/

		[Fact]
		public void CallPrivateMemberFromWithinContextOfClassShouldWork()
		{
			const string name = "test";
			var classWithPrivateMember = new ClassWithPrivateGetNameMethod(name);
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(
				name,
				_.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "GetName")
			);
		}

		[Fact]
		public void CallPrivateMemberFromOutsideContextOfClassShouldThrow()
		{
			const string name = "test";
			var classWithPrivateMember = new ClassWithPrivateGetNameMethod(name);
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Throws<MissingMemberException>(() =>
				_.CALL(context: null, target: classWithPrivateMember, member1: "GetName")
			);
		}

		[Fact]
		public void SettingPrivatePropertyFromWithinContextOfClassShouldWork()
		{
			const string name = "test";
			var classWithPrivateMember = new ClassWithPrivateNameProperty();
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			_.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Name");
			Assert.Equal(
				name,
				_.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "Name")
			);
		}

		[Fact]
		public void SettingPrivatePropertyFromOutsideContextOfClassShouldThrow()
		{
			const string name = "test";
			var classWithPrivateMember = new ClassWithPrivateNameProperty();
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Throws<MissingMemberException>(() =>
				_.SET(name, context: null, target: classWithPrivateMember, optionalMemberAccessor: "Name")
			);
		}

		private class ClassWithPrivateGetNameMethod
		{
			private readonly string _name;
			public ClassWithPrivateGetNameMethod(string name)
			{
				_name = name;
			}

			private string GetName()
			{
				return _name;
			}
		}

		private class ClassWithPrivateNameProperty // This is approximately how a VBScript property is translated (into getter and setter methods)
		{
			private object mname { get; set; }
			private object name()
			{
				return mname;
			}
			private void name(object strname)
			{
				mname = strname;
			}
		}

		[Fact]
		public void SettingPrivateIndexedPropertyFromWithinContextOfClassShouldWork()
		{
			const string name = "test";
			var i = new object();
			var classWithPrivateMember = new ClassWithPrivateIndexedProperty();
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			_.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i));
			Assert.Equal(
				name,
				_.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "Test", argumentProviderBuilder: _.ARGS.Val(i))
			);
		}

		[Fact]
		public void SettingPrivateIndexedPropertyFromOutsideContextOfClassShouldThrow()
		{
			const string name = "test";
			var i = new object();
			var classWithPrivateMember = new ClassWithPrivateIndexedProperty();
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Throws<MissingMethodException>(() =>
				_.SET(name, context: null, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i))
			);
		}

		// This is approximately how a VBScript indexed property is translated (into getter and setter methods) since C# only supports a single indexed property
		private class ClassWithPrivateIndexedProperty : TranslatedPropertyIReflectImplementation
		{
			private Dictionary<object, object> _values = new Dictionary<object, object>();

			[TranslatedProperty("Test")]
			private object test(object i)
			{
				return _values.ContainsKey(i) ? _values[i] : null;
			}

			[TranslatedProperty("Test")]
			private void test(object i, object value)
			{
				_values[i] = value;
			}
		}

		[Fact]
		public void SettingPublicIndexedPropertyFromWithinContextOfClassShouldWork()
		{
			const string name = "test";
			var i = new object();
			var classWithPrivateMember = new ClassWithPublicIndexedProperty();
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			_.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i));
			Assert.Equal(
				name,
				_.CALL(context: classWithPrivateMember, target: classWithPrivateMember, member1: "Test", argumentProviderBuilder: _.ARGS.Val(i))
			);
		}

		[Fact]
		public void SettingPublicIndexedPropertyFromOutsideContextOfClassShouldWork()
		{
			const string name = "test";
			var i = new object();
			var classWithPrivateMember = new ClassWithPublicIndexedProperty();
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			_.SET(name, context: classWithPrivateMember, target: classWithPrivateMember, optionalMemberAccessor: "Test", argumentProviderBuilder: _.ARGS.Val(i));
			Assert.Equal(
				name,
				_.CALL(context: null, target: classWithPrivateMember, member1: "Test", argumentProviderBuilder: _.ARGS.Val(i))
			);
		}

		// This is approximately how a VBScript indexed property is translated (into getter and setter methods) since C# only supports a single indexed property
		private class ClassWithPublicIndexedProperty : TranslatedPropertyIReflectImplementation
		{
			private Dictionary<object, object> _values = new Dictionary<object, object>();

			[TranslatedProperty("Test")]
			public object test(object i)
			{
				return _values.ContainsKey(i) ? _values[i] : null;
			}

			[TranslatedProperty("Test")]
			public void test(object i, object value)
			{
				_values[i] = value;
			}
		}

		[Fact]
		public void ByRefIndexArgumentOnPublicPropertySetterShouldAcceptUpdatesWhenCalledOverIReflect()
		{
			object i = "123";
			object value = "xyz";
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			_.SET(
				value,
				context: null,
				target: new ClassWithPublicIndexedPropertyThatHasByRefArguments(),
				optionalMemberAccessor: "Test",
				argumentProviderBuilder: _.ARGS.Ref(i, iUpdate => { i = iUpdate; })
			);
			Assert.Equal(i, "456");
		}

		private class ClassWithPublicIndexedPropertyThatHasByRefArguments : TranslatedPropertyIReflectImplementation
		{
			[TranslatedProperty("Test")]
			public void test(ref object i, ref object value)
			{
				i = "456";
			}
		}

		[Fact]
		public void ByRefIndexArgumentOnPublicPropertySetterShouldAcceptUpdatesWhenNotCalledOverIReflect()
		{
			object i = "123";
			object value = "xyz";
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			_.SET(
				value,
				context: null,
				target: new ClassWithPublicIndexedPropertyThatHasByRefArgumentsButThatIsNotCalledOverIReflect(),
				optionalMemberAccessor: "Test",
				argumentProviderBuilder: _.ARGS.Ref(i, iUpdate => { i = iUpdate; })
			);
			Assert.Equal(i, "456");
		}

		// Classes translated from VBScript that have indexed properties will be derived from TranslatedPropertyIReflectImplementation but we need to test the
		// logic when translated-from-VBScript code calls into not-translated-from-VBScript code as well (to ensure that the indexed arguments are ByRef-updated)
		private class ClassWithPublicIndexedPropertyThatHasByRefArgumentsButThatIsNotCalledOverIReflect
		{
			public void test(ref object i, ref object value)
			{
				i = "456";
			}
		}

		[Fact]
		public void CallingCLRMethodsThatHaveValueTypeParametersWorksWithReferenceTypes()
		{
			var recordset = new ADODB.Recordset();
			recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
			recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
			recordset.AddNew();
			recordset.Fields["name"].Value = "TestName";
			recordset.Update();

			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;

			object objField = _.CALL(
				context: null,
				target: recordset,
				members: new string[0],
				argumentProviderBuilder: _.ARGS.Val("name")
			);

			Assert.Equal(
				"TestName",
				_.CALL(
					context: null,
					target: this,
					member1: "MockMethodReturningInputString",
					argumentProviderBuilder: _.ARGS.Ref(objField, v => { objField = v; })
				)
			);
		}

		public string MockMethodReturningInputString(string input)
		{
			return input;
		}

		[Fact]
		public void NothingShouldBeReturnedForNullForPropertyOfComVisibleType()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			var value = _.CALL(context: null, target: new ClassWithComVisiblePropertyThatIsAlwaysNull(), member1: "Value");
			Assert.IsType<DispatchWrapper>(value);
			Assert.Null(((DispatchWrapper)value).WrappedObject);
		}

		[Fact]
		public void NothingShouldNotBeReturnedForNullForPropertyOfObjectType()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Null(
				_.CALL(context: null, target: new ClassWithObjectPropertyThatIsAlwaysNull(), member1: "Value")
			);
		}

		/// <summary>
		/// VBScript has its own ideas about what constitutes a value type, so it won't get Nothing from a null property value if the property's type is string (even though
		/// string is ComVisible and not "object" and not a .NET value type)
		/// </summary>
		[Fact]
		public void NothingShouldNotBeReturnedForNullForPropertyOfStringType()
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Null(
				_.CALL(context: null, target: new ClassWithObjectPropertyThatIsAlwaysNull(), member1: "Value")
			);
		}

		[ComVisible(true)]
		private class ClassWithComVisiblePropertyThatIsAlwaysNull
		{
			public ClassWithComVisiblePropertyThatIsAlwaysNull Value { get { return null; } }
		}

		[ComVisible(true)]
		private class ClassWithObjectPropertyThatIsAlwaysNull
		{
			public object Value { get { return null; } }
		}

		[Theory, MemberData("ZeroArgumentBracketSuccessData")]
		public void ZeroArgumentBracketSuccessCases(string description, object target, string[] memberAccessors, bool useBracketsWhereZeroArguments, object expectedResult)
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			var args = _.ARGS;
			if (useBracketsWhereZeroArguments)
				args = args.ForceBrackets();
			Assert.Equal(expectedResult, _.CALL(context: null, target: target, members: memberAccessors, argumentProvider: args.GetArgs()));
		}

		[Theory, MemberData("ZeroArgumentBracketFailData")]
		public void ZeroArgumentBracketFailCases(string description, object target, string[] memberAccessors, bool useBracketsWhereZeroArguments, Type exceptionType)
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			var args = _.ARGS;
			if (useBracketsWhereZeroArguments)
				args = args.ForceBrackets();
			Assert.Throws(exceptionType, () => _.CALL(context: null, target: target, members: memberAccessors, argumentProvider: args.GetArgs()));
		}

		public static IEnumerable<object[]> ZeroArgumentBracketSuccessData
		{
			get
			{
				var array = new object[] { 123 };
				yield return new object[] { "Array with no member accessors, properties or zero-argument brackets", array, new string[0], false, array };
				yield return new object[] { "String with no member accessors, properties or zero-argument brackets", "123", new string[0], false, "123" };

				var parameterLessDelegate = (Func<object>)(() => "delegate result");
				yield return new object[] { "Delegate with no member accessors, properties or zero-argument brackets", parameterLessDelegate, new string[0], false, parameterLessDelegate };
				yield return new object[] { "Delegate with no member accessors, properties WITH zero-argument brackets", parameterLessDelegate, new string[0], true, "delegate result" };

				yield return new object[] { "VBScript class property without brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "Name" }, false, "test" };
				yield return new object[] { "VBScript class property WITH brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "Name" }, true, "test" };
				yield return new object[] { "VBScript class function without brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "GetName" }, false, "test" };
				yield return new object[] { "VBScript class function WITH brackets", new ZeroArgumentBracketExampleClass("test"), new[] { "GetName" }, true, "test" };

				yield return new object[] { "COM component property without brackets", Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary")), new[] { "Count" }, false, 0 };
			}
		}

		public static IEnumerable<object[]> ZeroArgumentBracketFailData
		{
			get
			{
				yield return new object[] { "String with zero-argument brackets", "123", new string[0], true, typeof(TypeMismatchException) };
				yield return new object[] { "Array with zero-argument brackets", new object[] { 123 }, new string[0], true, typeof(SubscriptOutOfRangeException) };
				yield return new object[] { "COM component property with brackets", Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary")), new[] { "Count" }, true, typeof(IDispatchAccess.IDispatchAccessException) };
				yield return new object[] { "Delegate with a member accessors", (Func<object>)(() => "delegate result"), new[] { "Name" }, false, typeof(ArgumentException) };
			}
		}

		/// <summary>
		/// This will be used in tests that target "VBScript classes", meaning classes translated from VBScript into C# (as opposed to, say, COM components)
		/// </summary>
		private class ZeroArgumentBracketExampleClass
		{
			public ZeroArgumentBracketExampleClass(string name) { Name = name; }
			public string Name { get; private set; }
			public string GetName() { return Name; }
		}

		private class ImpressionOfTranslatedClassWithRewrittenPropertyName
		{
			/// <summary>
			/// This is a property that would have been rewritten from one named "Param" in VBScript (which is valid in VBScript but is a reserved
			/// keyword in C# and so may not appear in C# without some manipulation)
			/// </summary>
			public object rewritten_params { get { return "Success!"; } }
		}

		[Theory, MemberData("AcceptableEnumerableValueData")]
		public void AcceptableEnumerableValueCases(string description, object value, IEnumerable<object> expectedResults)
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Equal(_.ENUMERABLE(value).Cast<Object>(), expectedResults); // Cast to Object because we care about testing the contents, not the element type
		}

		[Theory, MemberData("UnacceptableEnumerableValueData")]
		public void UnacceptableEnumerableValueCases(string description, object value)
		{
			var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
			Assert.Throws<ObjectNotCollectionException>(() => _.ENUMERABLE(value));
		}

		public static IEnumerable<object[]> AcceptableEnumerableValueData
		{
			get
			{
				yield return new object[] { "An object array", new object[] { 1, 2 }, new object[] { 1, 2 } };

				dynamic dictionary = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
				dictionary.Add("key1", "value1");
				dictionary.Add("key2", "value2");
				yield return new object[] { "Scripting Dictionary COM component", dictionary, new object[] { "key1", "key2" } };

				var customIntEnumerable = new CustomEnumerable<int>(new[] { 1, 2, 3 });
				yield return new object[] { "Object with int values that has a valid GetEnumerator but does not implement IEnumerable", customIntEnumerable, new object[] { 1, 2, 3 } };

				var customIntEnumerableWithStructEnumerator = new CustomEnumerableWithStructEnumerator<int>(new[] { 1, 2, 3 });
				yield return new object[] { "Object with int values that has a valid GetEnumerator (that returns a struct) but does not implement IEnumerable", customIntEnumerableWithStructEnumerator, new object[] { 1, 2, 3 } };

				var customStringEnumerable = new CustomEnumerable<string>(new[] { "a", "b", "c" });
				yield return new object[] { "Object with string values that has a valid GetEnumerator but does not implement IEnumerable", customStringEnumerable, new object[] { "a", "b", "c" } };
			}
		}

		public static IEnumerable<object[]> UnacceptableEnumerableValueData
		{
			get
			{
				yield return new object[] { "Empty", null };
				yield return new object[] { "Null", DBNull.Value };
				yield return new object[] { "A string", "abc" }; // String ARE enumerable in C# but must not be treated so when mimicking VBScript
			}
		}

		private class ADOFieldObjectComparer : IEqualityComparer<object>
		{
			public new bool Equals(object x, object y)
			{
				if (x == null)
					throw new ArgumentNullException("x");
				if (y == null)
					throw new ArgumentNullException("y");
				var fieldX = x as ADODB.Field;
				if (fieldX == null)
					throw new ArgumentException("x is not an ADODB.Field");
				var fieldY = y as ADODB.Field;
				if (fieldY == null)
					throw new ArgumentException("y is not an ADODB.Field");
				return fieldX.Value == fieldY.Value;
			}

			public int GetHashCode(object obj)
			{
				return 0;
			}
		}

		private class PseudoFieldObjectComparer : IEqualityComparer<object>
		{
			public new bool Equals(object x, object y)
			{
				if (x == null)
					throw new ArgumentNullException("x");
				if (y == null)
					throw new ArgumentNullException("y");
				var fieldX = x as PseudoField;
				if (fieldX == null)
					throw new ArgumentException("x is not a PseudoField");
				var fieldY = y as PseudoField;
				if (fieldY == null)
					throw new ArgumentException("y is not a PseudoField");
				return (fieldX.value as string) == (fieldY.value as string);
			}

			public int GetHashCode(object obj)
			{
				return 0;
			}
		}

		private class PseudoRecordset
		{
			[IsDefault]
			public object fields(string fieldName)
			{
				return new PseudoField { value = "value:" + fieldName };
			}
		}

		private class PseudoField
		{
			[IsDefault]
			public object value { get; set; }
		}

		private sealed class CustomEnumerable<T>
		{
			private readonly IEnumerable<T> _values;
			public CustomEnumerable(IEnumerable<T> values)
			{
				if (values == null)
					throw new ArgumentNullException(nameof(values));
				_values = values;
			}
			public MyEnumerator<T> GetEnumerator() { return new MyEnumerator<T>(_values); }
		}

		private sealed class CustomEnumerableWithStructEnumerator<T>
		{
			private readonly IEnumerable<T> _values;
			public CustomEnumerableWithStructEnumerator(IEnumerable<T> values)
			{
				if (values == null)
					throw new ArgumentNullException(nameof(values));
				_values = values;
			}
			public MyStructEnumerator<T> GetEnumerator() { return new MyStructEnumerator<T>(new MyEnumerator<T>(_values)); }
		}

		private struct MyStructEnumerator<T>
		{
			private readonly MyEnumerator<T> _enumerator;
			public MyStructEnumerator(MyEnumerator<T> enumerator)
			{
				_enumerator = enumerator;
			}
			public T Current { get { return _enumerator.Current; } }
			public bool MoveNext() { return _enumerator.MoveNext(); }
			public void Reset() { _enumerator.Reset(); }
		}

		private sealed class MyEnumerator<T>
		{
			private readonly T[] _values;
			private int _index;
			public MyEnumerator(IEnumerable<T> values)
			{
				if (values == null)
					throw new ArgumentNullException(nameof(values));
				_values = values.ToArray();
				_index = -1;
			}
			public T Current
			{
				get
				{
					if (_index == -1)
						throw new InvalidOperationException("Enumeration has not started");
					return _values[_index];
				}
			}
			public bool MoveNext()
			{
				if (_index == _values.Length - 1)
					return false;

				_index++;
				return true;
			}
			public void Reset()
			{
				_index = -1;
			}
		}
	}
}
