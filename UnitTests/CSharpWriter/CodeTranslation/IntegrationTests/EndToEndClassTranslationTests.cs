using System;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
	public class EndToEndClassTranslationTests
	{
		/// <summary>
		/// When the tokens with the content "Property" had to be classified as a MayBeKeywordOrNameToken instead of a straight KeyWordToken, some logic had to
		/// be changed in the class parsing to account for it - this test exercises that work
		/// </summary>
		[Fact]
		public void EndProperty()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY GET Name
					END PROPERTY
				END CLASS
			";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}

					public object name()
					{
						return null;
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a class has a Class_Terminate method with at least one executable statement (ie. not empty and not just blank lines and comments), then it should
		/// implement IDisposable so that it's possible to instantiate it and tidy it up to simulate the way in which the deterministic VBScript interpreter
		/// calls Class_Terminate (as soon as it leaves scope, rather than when a garbage collector wants to deal with it). This isn't currently taken
		/// advantage of in the generated code (as of 2014-12-15) but it might be in the future.
		/// </summary>
		[Fact]
		public void ClassTerminateResultsInDisposableTranslatedClass()
		{
			var source = @"
				CLASS C1
					PUBLIC SUB Class_Terminate
						WScript.Echo ""Gone!""
					END SUB
				END CLASS
			";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1 : IDisposable
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					private bool _disposed1;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
						_disposed1 = false;
					}

					~c1()
					{
						try { Dispose2(false); }
						catch(Exception e)
						{
							try { _.SETERROR(e); } catch { }
						}
					}

					void IDisposable.Dispose()
					{
						Dispose2(true);
						GC.SuppressFinalize(this);
					}

					private void Dispose2(bool disposing)
					{
						if (_disposed1)
							return;
						if (disposing)
							class_terminate();
						_disposed1 = true;
					}

					public void class_terminate()
					{
						_.CALL(this, _env.wscript, ""Echo"", _.ARGS.Val(""Gone!""));
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a class has a Class_Initialize method with at least one executable statement (ie. not empty and not just blank lines and comments), then it should
		/// call this method in the constructor in the generated class. For strict compatibility with VBScript, any error is ignored and, while it will terminate
		/// execution of Class_Initialize, it will not prevent the calling code from continuing.
		/// </summary>
		[Fact]
		public void ClassInitializeResultsInConstructorCall()
		{
			var source = @"
				CLASS C1
					PUBLIC SUB Class_Initialize
						WScript.Echo ""Here!""
					END SUB
				END CLASS
			";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
						try { class_initialize(); }
						catch(Exception e)
						{
							_.SETERROR(e);
						}
					}

					public void class_initialize()
					{
						_.CALL(this, _env.wscript, ""Echo"", _.ARGS.Val(""Here!""));
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		// TODO: Need test here that has property that is set to a value in Class_Initialize
		[Fact]
		public void ClassInitializeCallHappensAfterFieldsSetToNull()
		{
			var source = @"
				CLASS C1
					PRIVATE mName
					PUBLIC SUB Class_Initialize
						mName = ""Test""
					END SUB
				END CLASS
			";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
						mname = null;
						try { class_initialize(); }
						catch(Exception e)
						{
							_.SETERROR(e);
						}
					}

					private object mname { get; set; }

					public void class_initialize()
					{
						mname = ""Test"";
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a GET property just returns a value then there's no need to define a return value, set it, then return that temporary reference (this is the
		/// same as for FUNCTION but not SUB, since SUB does not return a value)
		/// </summary>
		[Fact]
		public void PropertyGetterThatHasSingleLineReturnsIsTranslatedIntoSingleLineReturn()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY GET Name
						Name = ""C1""
					END PROPERTY
				END CLASS";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}
					public object name()
					{
						return ""C1"";
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// If a GET property is NOT just a single return-value-setting-statement then a temporary return reference is declared which is set (potentially
		/// multiple times, depending upon the getter's implementation) and returned at the end. The logic to determine whether a no-temporary-reference
		/// short cut may be applied is very simplistic and only applies to the simplest cases.
		/// </summary>
		[Fact]
		public void PropertyGetterThatHasMultipleLinesFollowsStandardFormat()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY GET Name
						WScript.Echo ""get_Name""
						Name = ""C1""
					END PROPERTY
				END CLASS";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}
					public object name()
					{
						object retVal1 = null;
						_.CALL(this, _env.wscript, ""Echo"", _.ARGS.Val(""get_Name""));
						retVal1 = ""C1"";
						return retVal1;
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// LET or SET properties would seem like they should error if they try to return a value (the same as a SUB would), but for some reason VBScript just
		/// ignores the sort-of-return-value setting (it evaluates the right-hand side of the statement but doesn't return anything and doesn't error)
		/// </summary>
		[Fact]
		public void NonGetPropertyIgnoresAnyReturnValueSetting()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Name(value)
						Name = ""C1""
					END PROPERTY
				END CLASS";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}
					public void name(ref object value)
					{
						_.VAL(""C1"");
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is similar to NonGetPropertyIgnoresAnyReturnValueSetting, where a LET property accessor includes a value-setting statement which appears to
		/// target the current property, but that value-setting statement specifies a SET. This means that the right-hand side of the statement must be of
		/// an object reference type. This is nothing to with whether the property accessor is a LET or SET, it is solely down to whether the value-setting
		/// statement begins with "SET" or not.
		/// </summary>
		[Fact]
		public void NonGetPropertyIgnoresAnyReturnValueSettingButSetSemanticsAreRespectedWhereSpecified()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Name(value)
						SET Name = ""C1""
					END PROPERTY
				END CLASS";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}
					public void name(ref object value)
					{
						_.OBJ(""C1"");
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// This is similar to the NonGetPropertyIgnoresAnyReturnValueSetting test except that if the left-hand side of a value-setting statement within a LET
		/// or SET property specifies the name of that property WITH brackets, then it will try to call itself (potentially infinite-looping, depending upon
		/// implementation)
		/// </summary>
		[Fact]
		public void NonGetPropertyCallsSelfIfBracketsAreSpecifiedAroundRecursivePropertyUpdate()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Name(value)
						Name() = ""C1""
					END PROPERTY
				END CLASS";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}
					public void name(ref object value)
					{
						_.SET(""C1"", this, this, ""Name"");
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}

		/// <summary>
		/// Indexed properties can not be directly represented in C# (well, one may be - the default indexed property - but it can't be explicitly named and things
		/// will fall apart if there need to be multiple indexed properties if this is the only mechanism used) so some extra logic is layered on; the properties
		/// are translated into functions and the parent class inherits TranslatedPropertyIReflectImplementation, which does some mapping work for calling code.
		/// </summary>
		[Fact]
		public void IndexedPropertiesNeedSpecialLoveAndCare()
		{
			var source = @"
				CLASS C1
					PUBLIC PROPERTY LET Blah(ByVal i, ByVal j, ByVal value)
					END PROPERTY
				END CLASS
			";
			var expected = @"
				[ComVisible(true)]
				[SourceClassName(""C1"")]
				public sealed class c1 : TranslatedPropertyIReflectImplementation
				{
					private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
					private readonly EnvironmentReferences _env;
					private readonly GlobalReferences _outer;
					public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
					{
						if (compatLayer == null)
							throw new ArgumentNullException(""compatLayer"");
						if (env == null)
							throw new ArgumentNullException(""env"");
						if (outer == null)
							throw new ArgumentNullException(""outer"");
						_ = compatLayer;
						_env = env;
						_outer = outer;
					}

					[TranslatedProperty(""Blah"")]
					public void blah(object i, object j, object value)
					{
					}
				}";
			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}
	}
}
