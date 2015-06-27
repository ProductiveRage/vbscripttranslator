using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndEraseTranslationTests
    {
        [Theory, MemberData("SuccessData")]
        public void SuccessCases(string description, string source, string[] expected)
        {
            Assert.Equal(expected, WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies));
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                yield return new object[] { "Empty ERASE is a runtime error", "ERASE", new[] { "throw new Exception(\"Wrong number of arguments: 'Erase' (line 1)\");" } };
                yield return new object[] { "Empty ERASE is a runtime error (with CALL keyword)", "CALL ERASE", new[] { "throw new Exception(\"Wrong number of arguments: 'Erase' (line 1)\");" } };

                yield return new object[] { "Simplest case: ERASE a", "ERASE a", new[] { "_.ERASE(_env.a, v1 => { _env.a = v1; });" } };
                yield return new object[] { "Simplest case: ERASE a (with CALL keyword)", "CALL ERASE(a)", new[] { "_.ERASE(_env.a, v1 => { _env.a = v1; });" } };

                // If the target is specified with arguments, then it must be an array where the arguments are indices. The non-by-ref ERASE method signature is used and validation of the
                // target (whether it's an array and whether the indices are valid) is handled at runtime.
                yield return new object[] { "Target with arguments: ERASE a(0)", "ERASE a(0)", new[] { "_.ERASE(_env.a, (Int16)0);" } };
                yield return new object[] { "Target with arguments: CALL ERASE(a(0)) (with CALL keyword)", "CALL ERASE(a(0))", new[] { "_.ERASE(_env.a, (Int16)0);" } };

                // "ERASE a()" is either a "Subscript out of range" or a "Type mismatch", depending upon whether "a" is an array or not - this needs to be decided at runtime. It does this
                // using the non-by-ref argument argument signature. This is the case where "a" is known to be a variable (whether explicitly declared or not, if "a" is known to be a
                // function then it's a different error case).
                yield return new object[] { "ERASE a()", "ERASE a()", new[] { "_.ERASE(_env.a);" } };

                yield return new object[] {
                    "Error if the target is known not to be a variable",
                    "ERASE a\nFUNCTION a\nEND FUNCTION",
                    new[] {
                        "var invalidEraseTarget1 = _.CALL(_outer, \"a\");",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");",
                        "public object a()",
                        "{",
                        "return null;",
                        "}"
                    }
                };
                yield return new object[] {
                    "Error if the target is known not to be a variable (takes precedence over other ERASE a() error case)",
                    "ERASE a()\nFUNCTION a\nEND FUNCTION",
                    new[] {
                        "var invalidEraseTarget1 = _.CALL(_outer, \"a\", _.ARGS.ForceBrackets());",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");",
                        "public object a()",
                        "{",
                        "return null;",
                        "}"
                    }
                };
                
                // Note: When the arguments are invalid, they are still evaluated and THEN the runtime error is raised. The references are not forced into value types (if they appear valid
                // at this point then the ERASE call must confirm at runtime that the target is an array), so the evaulation of some targets (eg. "a") will have no effect while others (eg.
                // "a.GetName()" may have side effects).
                yield return new object[] {
                    "Brackets around target (would be by-val => invalid)",
                    "ERASE (a)",
                    new[] {
                        "var invalidEraseTarget1 = _env.a;",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");"
                    }
                };
                yield return new object[] {
                    "Multiple targets",
                    "ERASE a, b",
                    new[] {
                        "var invalidEraseTarget1 = _env.a;",
                        "var invalidEraseTarget2 = _env.b;",
                        "throw new Exception(\"Wrong number of arguments: 'Erase' (line 1)\");"
                    }
                };
                yield return new object[] {
                    "Member access target",
                    "ERASE a.Name",
                    new[] {
                        "var invalidEraseTarget1 = _.CALL(_env.a, \"Name\");",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");"
                    }
                };
            }
        }

        [Fact]
        public void SingleTokenEraseTargetsRequireByRefAliasingIfTheTargetIsByRefArgumentOfTheContainingFunction()
        {
            var source = @"
                Function F1(a)
                    ERASE a
                End Function";
            var expected = @"
                public object f1(ref object a)
                {
                    object retVal1 = null;
                    object byrefalias2 = a;
                    try
                    {
                        _.ERASE(byrefalias2, v3 => { byrefalias2 = v3; });
                    }
                    finally { a = byrefalias2; }
                    return retVal1;
                }";
            Assert.Equal(
                expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Skip(1).Select(v => v.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}