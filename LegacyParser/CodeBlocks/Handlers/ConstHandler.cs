using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class ConstHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            if (!base.checkAtomTokenPattern(tokens, new string[] { "CONST" }, false))
                return null;

            tokens.RemoveAt(0); // Trim out the keyword before trying to extract the values being set
            var values = new List<ConstStatement.ConstValueInitialisation>();
            while (true)
            {
                if ((tokens.Count < 3) || !(tokens[0] is NameToken) || !(tokens[1] is OperatorToken) || (tokens[1].Content != "="))
                    throw new ArgumentException("Invalid input - encountered invalid CONST statement");

                // Note: The ConstValueInitialisation constructor will throw an exception if tokens[2] is not an acceptable value (it must
                // be a literal or one of a set of acceptable built-in values, such as Empty, Null or Nothing)
                values.Add(new ConstStatement.ConstValueInitialisation(
                    (NameToken)tokens[0],
                    tokens[2]
                ));

                // Remove the tokens we've consumed and any comma separators between values - exit if there is no separator indicating that
                // another value follows
                tokens.RemoveRange(0, 3);
                if (!tokens.Any() || !(tokens[0] is ArgumentSeparatorToken))
                    break;
                tokens.RemoveAt(0); // Remove the separator and try to process the next value
            }
            return new ConstStatement(values);
        }
    }
}
