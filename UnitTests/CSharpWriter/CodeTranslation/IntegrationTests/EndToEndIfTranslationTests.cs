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

        /// <summary>
        /// If both sides of a comparison are numeric constants, then there is no casting required since they are both known to be numeric at compile
        /// time (the test above illustrates what happens if one side is a numeric constant and the other isn't, this to show that that rule only
        /// holds when one side is a numeric constant and the other isn't, NOT just when one side is a numeric constant - the logic considers
        /// both sides of the operation)
        /// </summary>
        [Fact]
        public void NumericConstantComparedToAnotherNumericConstantDoesNotRequireAdditionalParsing()
        {
            var source = @"
			    If (1 = 2) Then
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.EQ(1, 2)))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// Even though built-in numeric constants (such as vbNull) are known to be numeric constants at compile time, they do NOT trigger the behaviour
        /// that is seen when a "real" numeric constant appears in a comparison (if the other side of the operation is a string, this is not forced to
        /// be parsed into a numeric value)
        /// </summary>
        [Fact]
        public void NumbericBuiltInValuesDoNotCountAsNumericConstantsWhenMakingComparisons()
        {
            var source = @"
			    If (""aa"" = vbNull) Then
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.EQ(\"aa\", _.Constants.vbNull)))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// Although boolean values are often treated as numbers in VBScript, if a comparison to a numeric constant is made then that boolean value
        /// must be translated into a number, just to be sure
        /// </summary>
        [Fact]
        public void BooleansDoNotGetTreatedAsNumericConstants()
        {
            var source = @"
			    If (False = 1) Then
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.EQ(_.NUM(_.Constants.False), 1)))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// It seems feasible that calculations that are constructed only of numeric constants would be treated as a numeric constant - eg. (1+0)
        /// could be identified at compile time as effectively being a numeric constant value. The VBScript interpreter does not do this, so if
        /// a string is compared to (1+0) then the string will not be parsed into a number before the runtime comparison method is called.
        /// </summary>
        [Fact]
        public void CalculationsThatCouldBeEvaluatedAsNumericConstantsAreNotTreatedAsNumericConstants()
        {
            var source = @"
			    If (""aa"" = (1+0)) Then
			    End If
            ";
            var expected = new[]
            {
                "if (_.IF(_.EQ(\"aa\", _.ADD(1, 0))))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// CalculationsThatCouldBeEvaluatedAsNumericConstantsAreNotTreatedAsNumericConstants may suggest that if numeric constants are bracketed
        /// away then they are no longer treated as numeric constants. This is not the case, though, unnecessary brackets are removed and so any
        /// string that is compared to that value must be parsed into a numeric value (or an error raised) before the comparison is made.
        /// </summary>
        [Fact]
        public void UnnecessaryBracketsAreUnrolledFromNumericConstants()
        {
            var source = @"
			    If (""aa"" = (1)) Then
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
    }
}
