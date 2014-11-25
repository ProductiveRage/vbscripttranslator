using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndIfTranslationTests
    {
        /// <summary>
        /// When running the parser against real content a silly mistake was found where an "ELSE" inside a comment would be treated as a
        /// real ELSE and result in an exception being raised when the real ELSE keyword was encountered
        /// </summary>
        [Fact]
        public void DoNotConsiderKeywordsInComments()
        {
            var source = @"
			    If True Then
				    'Else
			    Else
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.Constants.True))",
                "{",
                "  //Else",
                "}",
                "else",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// When testing against real content, another silly mistake was found where an inline comment would result in the parser getting confused
        /// </summary>
        [Fact]
        public void DoNotGetConfusedByCommentsInLineWithConditions()
        {
            var source = @"
			    If True Then 'Comment
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.Constants.True)) //Comment",
                "{",
                "}",
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This addresses a bug found in testing (relating to InlineCommentStatement detection in the IfBlockTranslator, which assumed that there would
        /// always be at least one statement within any conditional block)
        /// </summary>
        [Fact]
        public void EmptyIfBlocksDoNotCauseExceptions()
        {
            var source = @"
			    If True Then
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.Constants.True))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
