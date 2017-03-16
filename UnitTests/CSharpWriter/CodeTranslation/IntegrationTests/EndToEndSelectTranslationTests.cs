using System;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
	public class EndToEndSelectTranslationTests
	{
		/// <summary>
		/// This tests a fix made to select block translation - it was looking for token types based upon their content, rather than their type (so it was mistaking
		/// a StringToken whose content was a single comma characters as being an ArgumentSeparatorToken, if the type of the token is checked instead of its content
		/// then this sort of mistake will no longer occur)
		/// </summary>
		[Fact]
		public void AllowSpecialCharactersToBeUsedAsStringsInSelectCases()
		{
			var source = @"
				Select Case x
					Case ""(""
						WScript.Echo ""Open""
					Case "")""
						WScript.Echo ""Close""
					Case "",""
						WScript.Echo ""Split""
				End Select";

			var expected = @"
				if (_.IF(_.EQ(_env.x, ""("")))
				{
					_.CALL(this, _env.wscript, ""Echo"", _.ARGS.Val(""Open""));
				}
				else if (_.IF(_.EQ(_env.x, "")"")))
				{
					_.CALL(this, _env.wscript, ""Echo"", _.ARGS.Val(""Close""));
				}
				else if (_.IF(_.EQ(_env.x, "","")))
				{
					_.CALL(this, _env.wscript, ""Echo"", _.ARGS.Val(""Split""));
				}";

			Assert.Equal(
				expected.Replace(Environment.NewLine, "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
				WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
			);
		}
	}
}
