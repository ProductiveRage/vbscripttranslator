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
                "if (_.IF(VBScriptConstants.True))",
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
                "if (_.IF(VBScriptConstants.True)) //Comment",
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
                "if (_.IF(VBScriptConstants.True))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// There are a range of special cases if one side of a comparison is a numeric constant - the other side must be parsed into a numeric value for comparison
        /// (and a type mismatch error raised if this is not possible)
        /// </summary>
        public class NumericLiteralSpecialCases
        {
            /// <summary>
            /// If a comparison is made (such as an equality check or a greater-than comparison) and one side of the operation is a compile-time constant, then
            /// the other side of the opertion will be interpreted as a number and an error thrown if this fails (whereas, if the numeric side of the operation
            /// is a variable with a numeric value, a comparison to a value that is not parseable as a number would fail, rather than error - it only errors
            /// if one side is a number.
            /// </summary>
            [Fact]
            public void NonNegativeNumericConstantComparedToVariableRequiresTheVariableBeParsedToNumber()
            {
                var source = @"
			        If (i = 1) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_.NULLABLENUM(_env.i), (Int16)1)))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// Negative numbers are not considered to be numeric literals, they do not get the special treatment
            /// </summary>
            [Fact]
            public void NegativeNumericConstantsDoNotGetSpecialTreatment()
            {
                var source = @"
			        If (i = -1) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_env.i, (Int16)(-1))))",
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
                    "if (_.IF(_.EQ((Int16)1, (Int16)2)))",
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
            /// that is seen when a "real" numeric constant appears in a comparison (if the other side of the operation is a variable, this variable is not
            /// forced parsed into a numeric value)
            /// </summary>
            [Fact]
            public void NumericBuiltInValuesDoNotCountAsNumericConstantsWhenMakingComparisons()
            {
                var source = @"
			        If (i = vbNull) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_env.i, VBScriptConstants.vbNull)))",
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
                    "if (_.IF(_.EQ(_.NULLABLENUM(VBScriptConstants.False), (Int16)1)))",
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
            /// a variable is compared to (1+0) then the variable will not be parsed into a number before the runtime comparison method is called.
            /// </summary>
            [Fact]
            public void CalculationsThatCouldBeEvaluatedAsNumericConstantsAreNotTreatedAsNumericConstants()
            {
                var source = @"
			        If (i = (1+0)) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_env.i, _.ADD((Int16)1, (Int16)0))))",
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
            /// value that is compared to that value must be parsed into a numeric value (or an error raised) before the comparison is made.
            /// </summary>
            [Fact]
            public void UnnecessaryBracketsAreUnrolledFromNumericConstants()
            {
                var source = @"
			        If (i = (1)) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_.NULLABLENUM(_env.i), (Int16)1)))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// If a numeric value has TWO minus signs then you might think that they would cancel out and it be left as a positive numeric literal,
            /// to be processed as such. This is not the case since the minus signs prevent it from being categorised as a literal. The OperatorCombiner
            /// class will remove those double negatives but, if they are followed by a numeric constant, will wrap that constant in a CInt, CLng or CDbl
            /// call (depending upon the numeric value) so that it is obvious that it is not a literal (although, actually, the translator is clever enough
            /// to see when this might have happened and the CInt / CLng / CDbl call gets removed after it has done its job of preventing the values from
            /// being treated as a literal where it matters in the translation process.
            /// </summary>
            [Fact]
            public void DoubleNegativesAreRemovedButCanPreventsNumericLiteralSpecialBehaviour()
            {
                var source = @"
			        If (""12"" = --12) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(\"12\", _.STR((Int16)12))))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// This is extremely similar to DoubleNegativesAreRemovedButCanPreventsNumericLiteralSpecialBehaviour except that it is a single plus
            /// sign that is removed rather than double minus sign. Both have no effect on the value itself, but they do affect whether or not it
            /// is treated as a literal (it is not).
            /// </summary>
            [Fact]
            public void PlusSignBeforeNumberPreventsNumericLiteralSpecialBehaviour()
            {
                var source = @"
			        If (""12"" = +12) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(\"12\", _.STR((Int16)12))))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }
        }

        /// <summary>
        /// There are similar special cases if one side of a comparison is a string constant - the other side must be parsed into a string value for comparison
        /// </summary>
        public class StringLiteralSpecialCases
        {
            /// <summary>
            /// This is equivalent to the NumericConstantComparedToVariableRequiresTheVariableBeParsedToNumber test but it may seem harder to see how importance
            /// at a glance since there is no can-not-parse-into-number type mismatch that may be raised. The simplest example that illustrates it is to consider
            /// two variables; i = 1, j = "1". The comparison (i = j) returns false but (i = "1") returns true, even though j = "1".
            /// </summary>
            [Fact]
            public void StringConstantComparedToVariableResultsInTheOtherValueBeingTranslatedIntoString()
            {
                var source = @"
			        If (i = ""1"") Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_.STR(_env.i), \"1\")))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// If both sides of a comparison are string constants, then there is no casting required since they are both known to be string at compile time
            /// </summary>
            [Fact]
            public void StringConstantComparedToAnotherStringConstantDoesNotRequireAdditionalParsing()
            {
                var source = @"
			        If (""1"" = ""2"") Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(\"1\", \"2\")))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// If there is a case of where a numeric literal is compared to a string literal, the numeric literal rule takes precedence and the string
            /// value must be parsed into a number (or a type mismatch raised)
            /// </summary>
            [Fact]
            public void StringConstantWillBeParsedAsNumberIfComparedToNumericConstant()
            {
                var source = @"
			        If (""aa"" = 1) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_.NULLABLENUM(\"aa\"), (Int16)1)))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }

            /// <summary>
            /// There are no special cases around boolean literals, so if a boolean literal is compared to a string literal then the boolean is
            /// converted into a string, just like any other (non-numeric-literal) value will be
            /// </summary>
            [Fact]
            public void BooleanConstantComparedToStringConstantWillBeConsideredAsString()
            {
                var source = @"
			        If (""True"" = True) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(\"True\", _.STR(VBScriptConstants.True))))",
                    "{",
                    "}"
                };
                Assert.Equal(
                    expected.Select(s => s.Trim()).ToArray(),
                    WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
                );
            }
        }

        /// <summary>
        /// Boolean literals do NOT trigger special handling (it's a bug, as described at http://blogs.msdn.com/b/ericlippert/archive/2004/07/30/202432.aspx)
        /// </summary>
        public class BooeleanLiteralAbsentSpecialCases
        {
            [Fact]
            public void BooleanConstantComparedToVariableDoesNotResultInAnyFunnyBusiness()
            {
                var source = @"
			        If (i = True) Then
			        End If
                ";
                var expected = new[]
                {
                    "if (_.IF(_.EQ(_env.i, VBScriptConstants.True)))",
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
}
