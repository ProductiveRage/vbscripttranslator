using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.ContentBreaking;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser
{
    public static class Parser
    {
        public static IEnumerable<ICodeBlock> Parse(string scriptContent)
        {
            if (string.IsNullOrWhiteSpace(scriptContent))
                throw new ArgumentException("Null/blank scriptContent specified");

            // Break down content into String, Comment and UnprocessedContent tokens
            var tokens = StringBreaker.SegmentString(
                scriptContent.Replace("\r\n", "\n")
            );

            // Break down further into String, Comment, Atom and AbstractEndOfStatement tokens
            var atomTokens = new List<IToken>();
            foreach (var token in tokens)
            {
                if (token is UnprocessedContentToken)
                    atomTokens.AddRange(TokenBreaker.BreakUnprocessedToken((UnprocessedContentToken)token));
                else
                    atomTokens.Add(token);
            }

            // Translate these tokens into ICodeBlock implementations (representing code VBScript structures)
            string[] endSequenceMet;
            return (new CodeBlockHandler(null)).Process(atomTokens, out endSequenceMet);
        }
    }
}
