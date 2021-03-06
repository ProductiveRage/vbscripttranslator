﻿using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Handlers;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class ExitHandlerTests
    {
        [Fact]
        public void DoNotCrashIfReachEndOfContent()
        {
            var tokens = new List<IToken>
            {
                new KeyWordToken("Exit", 0),
                new KeyWordToken("function", 0)
            };
            (new ExitHandler()).Process(tokens);
        }
    }
}
