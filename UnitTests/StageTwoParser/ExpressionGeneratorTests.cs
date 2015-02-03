using CSharpWriter.CodeTranslation.Extensions;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using VBScriptTranslator.StageTwoParser.Tokens;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.StageTwoParser
{
    public class ExpressionGeneratorTests
    {
        [Fact]
        public void DirectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("Test", 0))
                    )
                },
                ExpressionGenerator.Generate(
                    new[] {
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

		[Fact]
        public void DirectFunctionCallWithNoArgumentsWithBrackets()
        {
            // Test()
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("Test", 0), CallExpressionSegment.ArgumentBracketPresenceOptions.Present)
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ObjectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // a.Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("a", 0), new NameToken("Test", 0) }
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new MemberAccessorToken(0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void NestedObjectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // a.b.Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("a", 0), new NameToken("b", 0), new NameToken("Test", 0) }
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new MemberAccessorToken(0),
                        new NameToken("b", 0),
                        new MemberAccessorToken(0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithOneArgument()
        {
            // Test(1)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("Test", 0) },
                            new[] { new NumericValueToken("1", 0) }
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken("1", 0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithTwoArguments()
        {
            // Test(1, 2)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("Test", 0) },
                            new[] { new NumericValueToken("1", 0) },
                            new[] { new NumericValueToken("2",0) }
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken("1", 0),
                        new ArgumentSeparatorToken(",", 0),
                        new NumericValueToken("2",0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void DirectFunctionCallWithTwoArgumentsOneIsNestedDirectionFunctionCallWithOneArgument()
        {
            // Test(Test2(1), 2)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new NameToken("Test", 0) },
                            EXP(
                                CALL(
                                    new[] { new NameToken("Test2", 0) },
                                    new[] { new NumericValueToken("1", 0) }
                                )
                            ),
                            EXP(CALL(new NumericValueToken("2",0)))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NameToken("Test2", 0),
                        new OpenBrace(0),
                        new NumericValueToken("1", 0),
                        new CloseBrace(0),
                        new ArgumentSeparatorToken(",", 0),
                        new NumericValueToken("2",0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ArrayElementFunctionCallWithNoArguments()
        {
            // a(0).Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a", 0) },
                                new[] { new NumericValueToken("0", 0) }
                            ),
                            CALL(
                                new[] { new NameToken("Test", 0) }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0),
                        new MemberAccessorToken(0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ObjectPropertyArrayElementFunctionCallWithNoArguments()
        {
            // a.b(0).Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a", 0), new NameToken("b", 0) },
                                new[] { new NumericValueToken("0", 0) }
                            ),
                            CALL(
                                new[] { new NameToken("Test", 0) }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new MemberAccessorToken(0),
                        new NameToken("b", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0),
                        new MemberAccessorToken(0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void ArrayElementNestedFunctionCallWithNoArguments()
        {
            // a(0).b.Test
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a", 0) },
                                new[] { new NumericValueToken("0", 0) }
                            ),
                            CALL(
                                new[] { new NameToken("b", 0), new NameToken("Test", 0) }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(new IToken[] {
                        new NameToken("a", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0),
                        new MemberAccessorToken(0),
                        new NameToken("b", 0),
                        new MemberAccessorToken(0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void JaggedArrayAccess()
        {
            // a(0)(1)
            Assert.Equal(new[]
                {
                    EXP(
                        CALLSET(
                            CALL(
                                new[] { new NameToken("a", 0) },
                                new[] { new NumericValueToken("0", 0) }
                            ),
                            CALLARGSONLY(
                                new[] { new NumericValueToken("1", 0) }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0),
                        new OpenBrace(0),
                        new NumericValueToken("1", 0),
                        new CloseBrace(0),
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Additional brackets will be applied around all operations to ensure that VBScript operator rules are always maintained (if the operators
        /// are all equivalent in terms of priority, terms will be bracketed from left-to-right, so a and b should be bracketed together)
        /// </summary>
        [Fact]
        public void AdditionWithThreeTerms()
        {
            // a + b + c
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            CALL(new NameToken("a", 0)),
                            OP(new OperatorToken("+", 0)),
                            CALL(new NameToken("b", 0))
                        ),
                        OP(new OperatorToken("+", 0)),
                        CALL(new NameToken("c", 0))
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new NameToken("b", 0),
                        new OperatorToken("+", 0),
                        new NameToken("c", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Multiplication should take precedence over addition so b and c should be bracketed together
        /// </summary>
        [Fact]
        public void AdditionAndMultiplicationWithThreeTerms()
        {
            // a + b * c
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a", 0)),
                        OP(new OperatorToken("+", 0)),
                        BR(
                            CALL(new NameToken("b", 0)),
                            OP(new OperatorToken("*", 0)),
                            CALL(new NameToken("c", 0))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new NameToken("b", 0),
                        new OperatorToken("*", 0),
                        new NameToken("c", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void AdditionAndMultiplicationWithThreeTermsWhereTheThirdTermIsAnArrayElement()
        {
            // a + b * c(0)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a", 0)),
                        OP(new OperatorToken("+", 0)),
                        BR(
                            CALL(new NameToken("b", 0)),
                            OP(new OperatorToken("*", 0)),
                            CALL(
                                new[] { new NameToken("c", 0) },
                                new[] { new NumericValueToken("0", 0) }
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new NameToken("b", 0),
                        new OperatorToken("*", 0),
                        new NameToken("c", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This will try to ensure that the bracket around the array access doesn't interfere with the formatting of the fourth term
        /// </summary>
        [Fact]
        public void AdditionAndMultiplicationAndAdditionWithFourTermsWhereTheThirdTermIsAnArrayElement()
        {
            // a + b * c(0) + d
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            CALL(new NameToken("a", 0)),
                            OP(new OperatorToken("+", 0)),
                            BR(
                                CALL(new NameToken("b", 0)),
                                OP(new OperatorToken("*", 0)),
                                CALL(
                                    new[] { new NameToken("c", 0) },
                                    new[] { new NumericValueToken("0", 0) }
                                )
                            )
                        ),
                        OP(new OperatorToken("+", 0)),
                        CALL(new NameToken("d", 0))
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new NameToken("b", 0),
                        new OperatorToken("*", 0),
                        new NameToken("c", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0),
                        new OperatorToken("+", 0),
                        new NameToken("d", 0),
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// If an operation is already bracketed then additional brackets should not be added around the operation, they would be unnecessary
        /// </summary>
        [Fact]
        public void AlreadyBracketedOperationsShouldNotGetUnnecessaryBracketing()
        {
            // a + (b * c)
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a", 0)),
                        OP(new OperatorToken("+", 0)),
                        BR(
                            CALL(new NameToken("b", 0)),
                            OP(new OperatorToken("*", 0)),
                            CALL(new NameToken("c", 0))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new OpenBrace(0),
                        new NameToken("b", 0),
                        new OperatorToken("*", 0),
                        new NameToken("c", 0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void AlreadyBracketedOperationsShouldNotGetUnnecessaryBracketingIfTheyAppearInTheMiddleOfTheExpression()
        {
            // a + (b * c) + d
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            CALL(new NameToken("a", 0)),
                            OP(new OperatorToken("+", 0)),
                            BR(
                                CALL(new NameToken("b", 0)),
                                OP(new OperatorToken("*", 0)),
                                CALL(new NameToken("c", 0))
                            )
                        ),
                        OP(new OperatorToken("+", 0)),
                        CALL(new NameToken("d", 0))
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new OpenBrace(0),
                        new NameToken("b", 0),
                        new OperatorToken("*", 0),
                        new NameToken("c", 0),
                        new CloseBrace(0),
                        new OperatorToken("+", 0),
                        new NameToken("d", 0),
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// Arithmetic operations should take precedence over comparisons so b and c should be bracketed together
        /// </summary>
        [Fact]
        public void AdditionAndEqualityComparisonWithThreeTerms()
        {
            // a = b + c
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a", 0)),
                        OP(new ComparisonOperatorToken("=", 0)),
                        BR(
                            CALL(new NameToken("b", 0)),
                            OP(new OperatorToken("+", 0)),
                            CALL(new NameToken("c", 0))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new ComparisonOperatorToken("=", 0),
                        new NameToken("b", 0),
                        new OperatorToken("+", 0),
                        new NameToken("c", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This covers an array of different types of expression
        /// </summary>
        [Fact]
        public void TestArrayAccessObjectAccessMethodArgumentsMixedArithmeticAndComparisonOperations()
        {
            // a + b * c.d(Test(0), 1) + e = f
            Assert.Equal(new[]
                {
                    EXP(
                        BR(
                            BR(
                                CALL(new NameToken("a", 0)),
                                OP(new OperatorToken("+", 0)),
                                BR(
                                    CALL(new NameToken("b", 0)),
                                    OP(new OperatorToken("*", 0)),
                                    CALL(
                                        new[] { new NameToken("c", 0), new NameToken("d", 0) },
                                        EXP(
                                            CALL(
                                                new[] { new NameToken("Test", 0) },
                                                new[] { new NumericValueToken("0", 0) }
                                            )
                                        ),
                                        EXP(CALL(new NumericValueToken("1", 0)))
                                    )
                                )
                            ),
                            OP(new OperatorToken("+", 0)),
                            CALL(new NameToken("e", 0))
                        ),
                        OP(new ComparisonOperatorToken("=", 0)),
                        CALL(new NameToken("f", 0))
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("+", 0),
                        new NameToken("b", 0),
                        new OperatorToken("*", 0),
                        new NameToken("c", 0),
                        new MemberAccessorToken(0),
                        new NameToken("d", 0),
                        new OpenBrace(0),
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new NumericValueToken("0", 0),
                        new CloseBrace(0),
                        new ArgumentSeparatorToken(",", 0),
                        new NumericValueToken("1", 0),
                        new CloseBrace(0),
                        new OperatorToken("+", 0),
                        new NameToken("e", 0),
                        new ComparisonOperatorToken("=", 0),
                        new NameToken("f", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// To make it clear that the "-" is a one-sided operation (a negation, not a subtraction), it should be bracketed
        /// </summary>
        [Fact]
        public void NegatedTermsShouldBeBracketed()
        {
            // a * -b
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a", 0)),
                        OP(new OperatorToken("*", 0)),
                        BR(
                            OP(new OperatorToken("-", 0)),
                            CALL(new NameToken("b", 0))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new OperatorToken("*", 0),
                        new OperatorToken("-", 0),
                        new NameToken("b", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This is the boolean equivalent of NegatedTermsShouldBeBracketed
        /// </summary>
        [Fact]
        public void LogicalInversionsTermsShouldBeBracketed()
        {
            // a AND NOT b
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(new NameToken("a", 0)),
                        OP(new LogicalOperatorToken("AND", 0)),
                        BR(
                            OP(new LogicalOperatorToken("NOT", 0)),
                            CALL(new NameToken("b", 0))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new NameToken("a", 0),
                        new LogicalOperatorToken("AND", 0),
                        new LogicalOperatorToken("NOT", 0),
                        new NameToken("b", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// This indicates different precedence that is applied to a NOT operation depending upon content, as compared to the test
        /// LogicalInversionsTermsShouldBeBracketed
        /// </summary>
        [Fact]
        public void NegationOperationHasLessPrecendenceThanComparsionOperations()
        {
            // NOT a IS Nothing
            Assert.Equal(new[]
                {
                    EXP(
                        OP(new LogicalOperatorToken("NOT", 0)),
                        BR(
                            CALL(new NameToken("a", 0)),
                            OP(new ComparisonOperatorToken("IS", 0)),
                            new BuiltInValueExpressionSegment(new BuiltInValueToken("Nothing", 0))
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new LogicalOperatorToken("NOT", 0),
                        new NameToken("a", 0),
                        new ComparisonOperatorToken("IS", 0),
                        new BuiltInValueToken("Nothing", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        [Fact]
        public void NewInstanceRequestsShouldNotBeConfusedWithCallExpressions()
        {
            // new Test
            Assert.Equal(new[]
                {
                    EXP(
                        NEW("Test", 0)
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new KeyWordToken("new", 0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }

        /// <summary>
        /// If a function (or property) argument is wrapped in brackets then it should be passed ByVal even when otherwise it would be passed ByRef.
        /// This means that brackets can have special significance and should not be removed, even from places where they would have significance or
        /// meaning in C#.
        /// </summary>
        [Fact]
        public void BracketsShouldNotBeRemovedFromSingleArgumentCallStatements()
        {
            // CALL Test((a))
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new NameToken("Test", 0),
                            EXP(
                                BR(CALL(new NameToken("a", 0)))
                            )
                        )
                    )
                },
                ExpressionGenerator.Generate(
                        new IToken[] {
                        new NameToken("Test", 0),
                        new OpenBrace(0),
                        new OpenBrace(0),
                        new NameToken("a", 0),
                        new CloseBrace(0),
                        new CloseBrace(0)
                    },
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }


        [Fact]
        public void ObjectFunctionCallWithNoArgumentsAndNoBracketsThatReliesUponDirectedWithReference()
        {
            // ".Test" within "WITH a"
            Assert.Equal(new[]
                {
                    EXP(
                        CALL(
                            new[] { new DoNotRenameNameToken("a", 0), new NameToken("Test", 0) }
                        )
                    )
                },
                ExpressionGenerator.Generate(
                    new IToken[] {
                        new MemberAccessorToken(0),
                        new NameToken("Test", 0)
                    },
                    directedWithReferenceIfAny: new DoNotRenameNameToken("a", 0),
                    warningLogger: warning => { }
                ),
                new ExpressionSetComparer()
            );
        }


        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        private static BracketedExpressionSegment BR(IEnumerable<IExpressionSegment> segments)
        {
			return new BracketedExpressionSegment(segments);
        }

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        private static BracketedExpressionSegment BR(params IExpressionSegment[] segments)
        {
            return new BracketedExpressionSegment((IEnumerable<IExpressionSegment>)segments);
        }

        private static CallSetExpressionSegment CALLSET(params IExpressionSegment[] segments)
        {
            return new CallSetExpressionSegment(segments.Cast<CallSetItemExpressionSegment>());
        }

		/// <summary>
		/// Create an CallExpressionSegment from member access tokens and argument expressions (the zeroArgBrackets is only considered if arguments is an empty set,
		/// if arguments is empty and zeroArgBrackets is null then a Absent will be used as a default)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments, CallExpressionSegment.ArgumentBracketPresenceOptions? zeroArgBrackets)
		{
            if ((memberAccessTokens.Count() == 1) && !arguments.Any())
            {
                if (memberAccessTokens.Single() is NumericValueToken)
                    return new NumericValueExpressionSegment(memberAccessTokens.Single() as NumericValueToken);
                if (memberAccessTokens.Single() is StringToken)
                    return new StringValueExpressionSegment(memberAccessTokens.Single() as StringToken);
            }

			CallExpressionSegment.ArgumentBracketPresenceOptions? argBrackets;
			if (arguments.Any())
				argBrackets = null;
			else if (zeroArgBrackets == null)
				argBrackets = CallExpressionSegment.ArgumentBracketPresenceOptions.Absent;
			else
				argBrackets = zeroArgBrackets;

            if (memberAccessTokens.Any())
            {
			    return new CallExpressionSegment(
                    memberAccessTokens,
                    arguments,
				    argBrackets
                );
            }
			return new CallSetItemExpressionSegment(
                memberAccessTokens,
                arguments,
				argBrackets
            );
        }

		/// <summary>
		/// Create a CallExpressionSegment from member access tokens and argument expressions (the zeroArgBrackets is only considered if arguments is an empty set,
		/// if arguments is empty and zeroArgBrackets is null then a Absent will be used as a default)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, CallExpressionSegment.ArgumentBracketPresenceOptions? zeroArgBrackets, params Expression[] arguments)
		{
			return CALL(memberAccessTokens, (IEnumerable<Expression>)arguments, zeroArgBrackets);
		}

		/// <summary>
		/// Create a CallExpressionSegment from member access tokens and argument expressions (applying the default logic for ArgumentBracketPresenceOptions; null
		/// if there are arguments and Absent otherwise)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params Expression[] arguments)
		{
			return CALL(memberAccessTokens, (IEnumerable<Expression>)arguments, null);
		}

		private static IExpressionSegment CALLARGSONLY(params IEnumerable<IToken>[] arguments)
		{
			return CALL(new IToken[0], arguments);
		}

		/// <summary>
		/// Create a CallExpressionSegment from a single member access token and argument expressions (applying the default logic for ArgumentBracketPresenceOptions;
		/// null if there are arguments and Absent otherwise)
		/// </summary>
		private static IExpressionSegment CALL(IToken memberAccessToken, params Expression[] arguments)
		{
			return CALL(new[] { memberAccessToken }, arguments);
		}

		/// <summary>
		/// Create a CallExpressionSegment from a single member access token with no argument expressions and an explicit ArgumentBracketPresenceOptions value
		/// </summary>
		private static IExpressionSegment CALL(IToken memberAccessToken, CallExpressionSegment.ArgumentBracketPresenceOptions zeroArgBrackets)
		{
			return CALL(new[] { memberAccessToken }, new Expression[0], zeroArgBrackets);
		}

		/// <summary>
		/// Create a CallExpressionSegment from a single member access token and argument expressions expressed as token sets (applying the default logic for
		/// ArgumentBracketPresenceOptions; null if there are arguments and Absent otherwise)
        /// </summary>
        private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params IEnumerable<IToken>[] arguments)
        {
            if ((memberAccessTokens.Count() == 1) && !arguments.Any())
            {
                if (memberAccessTokens.Single() is NumericValueToken)
                    return new NumericValueExpressionSegment(memberAccessTokens.Single() as NumericValueToken);
                if (memberAccessTokens.Single() is StringToken)
                    return new StringValueExpressionSegment(memberAccessTokens.Single() as StringToken);
            }
            return CALL(
                memberAccessTokens,
                arguments.Select(a => new Expression(new[] { CALL(a) })),
				null
            );
        }

        private static NewInstanceExpressionSegment NEW(string className, int lineIndex)
        {
			return new NewInstanceExpressionSegment(new NameToken(className, lineIndex));
        }

        private static OperationExpressionSegment OP(OperatorToken token)
        {
            return new OperationExpressionSegment(token);
        }

        /// <summary>
        /// Create an Expression from multiple ExpressionSegments
        /// </summary>
        private static Expression EXP(params IExpressionSegment[] segments)
        {
            return new Expression(segments);
        }

        /// <summary>
        /// Create an Expression from a single ExpressionSegment
        /// </summary>
        private static Expression EXP(IExpressionSegment segment)
        {
            return EXP(new[] { segment });
        }
    }
}
