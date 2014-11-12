using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Handlers;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class DimHandlerTests
    {
        /// <summary>
        /// There was an issue where the argument separator tokens weren't being removed from DIM statements for multiple variables - this is the
        /// fail-before-fixing test for that issue
        /// </summary>
        [Fact]
        public void VariableSeparatorsAreCorrectlyRemovedAsProcessedContent()
        {
            var tokens = new List<IToken>
            {
                new KeyWordToken("Dim", 0),
                new NameToken("i", 0),
                new ArgumentSeparatorToken(",", 0),
                new NameToken("j", 0),
                new ArgumentSeparatorToken(",", 0),
                new NameToken("k", 0),
                new ArgumentSeparatorToken(",", 0),
                new NameToken("l", 0)
            };
            (new DimHandler()).Process(tokens);
            Assert.Equal(0, tokens.Count);
        }
    }
}
