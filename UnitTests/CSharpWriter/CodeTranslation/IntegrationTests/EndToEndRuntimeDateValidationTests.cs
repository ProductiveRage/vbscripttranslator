using System;
using System.Linq;
using VBScriptTranslator.CSharpWriter;
using VBScriptTranslator.CSharpWriter.CodeTranslation.BlockTranslators;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
	public class EndToEndRuntimeDateValidationTests
	{
		/// <summary>
		/// If date literals are present in the source that need to be validated when the translated program is run (but before it does any other work), then extra code must be generated
		/// </summary>
		[Fact]
		public void RuntimeDateLiteralPresent()
		{
			var source = "If (a = #29 May 2015#) Then\nEnd If";
			var expected = @"
				using System;
				using System.Collections;
				using System.Collections.ObjectModel;
				using System.Runtime.InteropServices;
				using VBScriptTranslator.RuntimeSupport;
				using VBScriptTranslator.RuntimeSupport.Attributes;
				using VBScriptTranslator.RuntimeSupport.Exceptions;
				using VBScriptTranslator.RuntimeSupport.Compat;

				namespace TranslatedProgram
				{
					public class Runner
					{
						private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
						public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
						{
							if (compatLayer == null)
								throw new ArgumentNullException(""compatLayer"");
							_ = compatLayer;
						}

						public GlobalReferences Go()
						{
							return Go(new EnvironmentReferences());
						}
						public GlobalReferences Go(EnvironmentReferences env)
						{
							if (env == null)
								throw new ArgumentNullException(""env"");

							var _env = env;
							var _outer = new GlobalReferences(_, _env);
							RuntimeDateLiteralValidator.ValidateAgainstCurrentCulture(_);

							if (_.IF(_.EQ(_.NullableDATE(_env.a), _.DateLiteralParser.Parse(""29 May 2015""))))
							{
							}

							return _outer;
						}

						private static class RuntimeDateLiteralValidator
						{
							private static readonly ReadOnlyCollection<Tuple<string, int[]>> _literalsToValidate =
							new ReadOnlyCollection<Tuple<string, int[]>>(new[] {
								Tuple.Create(""29 May 2015"", new[] { 1 })
							});

							public static void ValidateAgainstCurrentCulture(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
							{
								if (compatLayer == null)
									throw new ArgumentNullException(""compatLayer"");
								foreach (var dateLiteralValueAndLineNumbers in _literalsToValidate)
								{
									try { compatLayer.DateLiteralParser.Parse(dateLiteralValueAndLineNumbers.Item1); }
									catch
									{
										throw new SyntaxError(string.Format(
											""Invalid date literal #{0}# on line{1} {2}"",
											dateLiteralValueAndLineNumbers.Item1,
											(dateLiteralValueAndLineNumbers.Item2.Length == 1) ? """" : ""s"",
											string.Join<int>("", "", dateLiteralValueAndLineNumbers.Item2)
										));
									}
								}
							}
						}

						public class GlobalReferences
						{
							private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
							private readonly GlobalReferences _outer;
							private readonly EnvironmentReferences _env;
							public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env)
							{
								if (compatLayer == null)
									throw new ArgumentNullException(""compatLayer"");
								if (env == null)
									throw new ArgumentNullException(""env"");
								_ = compatLayer;
								_env = env;
								_outer = this;
							}
						}

						public class EnvironmentReferences
						{
							public object a { get; set; }
						}
					}
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				DefaultTranslator.Translate(source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable).Select(s => s.Content.Trim()).Where(s => s != "").ToArray()
			);
		}

		/// <summary>
		/// If the only date literals can be safely validated at translation time and will not vary by culture, then there is no need to emit the ValidateAgainstCurrentCulture code
		/// </summary>
		[Fact]
		public void NoRuntimeDateLiteralPresent()
		{
			var source = "If (a = #29 5 2015#) Then\nEnd If";
			var expected = @"
				using System;
				using System.Collections;
				using System.Runtime.InteropServices;
				using VBScriptTranslator.RuntimeSupport;
				using VBScriptTranslator.RuntimeSupport.Attributes;
				using VBScriptTranslator.RuntimeSupport.Exceptions;
				using VBScriptTranslator.RuntimeSupport.Compat;

				namespace TranslatedProgram
				{
					public class Runner
					{
						private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
						public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
						{
							if (compatLayer == null)
								throw new ArgumentNullException(""compatLayer"");
							_ = compatLayer;
						}

						public GlobalReferences Go()
						{
							return Go(new EnvironmentReferences());
						}
						public GlobalReferences Go(EnvironmentReferences env)
						{
							if (env == null)
								throw new ArgumentNullException(""env"");

							var _env = env;
							var _outer = new GlobalReferences(_, _env);

							if (_.IF(_.EQ(_.NullableDATE(_env.a), _.DateLiteralParser.Parse(""29 5 2015""))))
							{
							}

							return _outer;
						}

						public class GlobalReferences
						{
							private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
							private readonly GlobalReferences _outer;
							private readonly EnvironmentReferences _env;
							public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env)
							{
								if (compatLayer == null)
									throw new ArgumentNullException(""compatLayer"");
								if (env == null)
									throw new ArgumentNullException(""env"");
								_ = compatLayer;
								_env = env;
								_outer = this;
							}
						}

						public class EnvironmentReferences
						{
							public object a { get; set; }
						}
					}
				}";

			Assert.Equal(
				expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
				DefaultTranslator.Translate(source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable).Select(s => s.Content.Trim()).Where(s => s != "").ToArray()
			);
		}
	}
}
