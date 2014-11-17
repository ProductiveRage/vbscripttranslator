using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public class EndToEndDoTranslationTests
    {
        [Fact]
        public void SimpleDoWhile()
        {
            var source = @"
                DO WHILE i > 10
                LOOP
            ";
            var expected = new[]
            {
                "do while (_.IF(_.GT(_env.i, 10)))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void SimpleDoUntil()
        {
            var source = @"
                DO UNTIL i > 10
                LOOP
            ";
            var expected = new[]
            {
                "do while (!_.IF(_.GT(_env.i, 10)))",
                "{",
                "}"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void SimpleDoLoopWhile()
        {
            var source = @"
                DO
                LOOP WHILE i > 10
            ";
            var expected = new[]
            {
                "do",
                "{",
                "} while (_.IF(_.GT(_env.i, 10)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void SimpleDoLoopUntil()
        {
            var source = @"
                DO
                LOOP UNTIL i > 10
            ";
            var expected = new[]
            {
                "do",
                "{",
                "} while (!_.IF(_.GT(_env.i, 10)));"
            };
            Assert.Equal(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [Fact]
        public void DoLoopWithoutTerminationCondition()
        {
            var source = @"
                DO
                LOOP
            ";
            var expected = new[]
            {
                "while (true)",
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
