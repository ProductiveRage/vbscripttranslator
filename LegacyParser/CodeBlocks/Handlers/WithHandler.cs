using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class WithHandler : AbstractBlockHandler
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

            if (!base.checkAtomTokenPattern(tokens, new string[] { "WITH" }, matchCase: false))
                return null;
            if (tokens.Count < 4)
                throw new ArgumentException("Insufficient tokens - invalid");

            // The WITH target will normally be a NameToken - eg. "WITH a" - but may also be wrapped in brackets - eg. "WITH (a)" - or may even
            // be another redirected reference (from an ancester WITH) - eg. "WITH .Item". We'll use the StatementHandler to determine what
            // content is part of the WITH target, but we don't directly require the returned Statement - we just needs its tokens (to
            // generate an Expression for the WithBlock).
            var token = base.getToken(tokens, offset: 1, allowedTokenTypes: new Type[] { typeof(OpenBrace), typeof(MemberAccessorOrDecimalPointToken), typeof(NameToken) });
            var targetTokensSource = tokens.Skip(1).ToList();
            var numberOfItemsInTargetTokensSource = targetTokensSource.Count;
            var target = new StatementHandler().Process(targetTokensSource) as Statement;
            if (target == null)
                throw new ArgumentException("The WITH target must be parseable as a (non-value-setting) statement");
            else if (target.CallPrefix == Statement.CallPrefixOptions.Present)
                throw new ArgumentException("The WITH target must be parseable as a statement without a CALL prefix");
            var numberOfItemsProcessedInTarget = numberOfItemsInTargetTokensSource - targetTokensSource.Count;
            tokens.RemoveRange(0, 1 + numberOfItemsProcessedInTarget); // Remove the "WITH" plus the tokens in the target reference

            // Get block content
            string[] endSequenceMet;
            var endSequences = new List<string[]>
            {
                new string[] { "END", "WITH" }
            };
            var codeBlockHandler = new CodeBlockHandler(endSequences);
            var blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find end sequence!");

            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if (tokens.Count > 0)
            {
                if (tokens[0] is AbstractEndOfStatementToken)
                    tokens.RemoveAt(0);
                else
                    throw new Exception("EndOfStatementToken missing after END WITH");
            }

            // Return code block instance
            return new WithBlock(new Expression(target.Tokens), blockContent);
        }
    }
}
