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

        /// <summary>
        /// If a comparison is made (such as an equality check or a greater-than comparison) and one side of the operation is a compile-time constant, then
        /// the other side of the opertion will be interpreted as a number and an error thrown if this fails (whereas, if the numeric side of the operation
        /// is a variable with a numeric value, a comparison to a string that is not parseable as a number would fail, rather than error - it only errors
        /// if one side is a number.
        /// </summary>
        [Fact]
        public void NumericConstantComparedToStringConstantRequiresTheStringBeParsedToNumber()
        {
            var source = @"
			    If (""aa"" = 1) Then
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.EQ(_.NUM(\"aa\"), 1)))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        // TODO: Test 1 against 2 (no cast required, as both already numeric)
        // TODO: Test "aa" against vbNull (no cast required - builtin values are not treated as numeric constants, so vbNull does not require the string be cast)
        // TODO: Test False against 1 (cast IS required - builtin value false is not treated as a numeric constant and so must be cast)
        // TODO: Test "aa" against (1+0) (no cast is required since the RHS is an operation, not a numeric constant)
        // TODO: Test "aa" against (10 (cast IS required since the redundant brackets are ignored)
    }
}
