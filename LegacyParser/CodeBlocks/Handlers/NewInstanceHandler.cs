using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class NewInstanceHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            if (!base.checkAtomTokenPattern(tokens, new string[] { "NEW" }, false))
                return null;
            if (tokens.Count < 2)
                throw new ArgumentException("Insufficient tokens - invalid");

            tokens.RemoveAt(0);
            var classNameToken = tokens[0] as NameToken;
            if (classNameToken == null)
                throw new ArgumentException("Token after the \"NEW\" keyword must be a NameToken");
            tokens.RemoveAt(0);
            if (tokens.Count > 0)
            {
                var endOfLineToken = tokens[0] as AbstractEndOfStatementToken;
                if (endOfLineToken == null)
                    throw new ArgumentException("The class name of a new-instance statement must be followed by an end-of-statement token");
                tokens.RemoveAt(0);
            }

            return new Statement(
                new IToken[]
                {
                    new KeyWordToken("new"),
                    classNameToken
                },
                Statement.CallPrefixOptions.Absent
            );
        }
    }
}
