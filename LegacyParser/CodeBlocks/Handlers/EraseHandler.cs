using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class EraseHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            bool includesCallKeyword;
            if (base.checkAtomTokenPattern(tokens, new[] { "CALL", "ERASE" }, false))
                includesCallKeyword = true;
            else if (base.checkAtomTokenPattern(tokens, new[] { "ERASE" }, false))
                includesCallKeyword = false;
            else
                return null;

            // Note: getEntryList will return a list of lists of tokens, where there will be multiple lists if there are multiple comma-separated expressions. Each
            // of these token sets needs to be wrapped in an Expression to initialise an EraseStatement. VBScript only works with a single target that is an array,
            // if there are zero or multiple targets then it will fail at runtime (but it's not a compile error so it's not an error case here).
            // - If the CALL keyword is present, then valid VBScript will have wrapped the argument(s) in brackets which will need stripping and then the arguments
            //   re-parsing. The getEntryList function will throw an exception if arguments are mismatched, indicating invalid VBScript. The translation process
            //   relies upon the source code not having any VBScript compile errors (in some places it tries to be helpful with explaining where there would be
            //   compile failures in the VBScript interpreter and in other places - like here - it pretty much just assume valid content.
            var keywordLineIndex = tokens[0].LineIndex;
            var targetExpressionsTokenSets = base.getEntryList(tokens, offset: (includesCallKeyword ? 2 : 1), endMarker: new EndOfStatementNewLineToken(tokens.First().LineIndex));
            if (includesCallKeyword)
            {
                if (targetExpressionsTokenSets.Count != 1)
                    throw new Exception("Expected only a single argument token set to have been extracted when CALL is present, since brackets should have wrapped all argument(s) in valid VBScript");
                var argumentTokens = targetExpressionsTokenSets[0];
                var terminator = new EndOfStatementNewLineToken(argumentTokens.Last().LineIndex);
                targetExpressionsTokenSets = base.getEntryList(
                    argumentTokens.Skip(1).Take(argumentTokens.Count - 2).Concat(new[] { terminator }),
                    0,
                    terminator
                );
            }
            var targetDetails = targetExpressionsTokenSets.Select(targetTokens => GetTargetExpressionDetailsWithNumberOfTokensConsumed(targetTokens));
            var numberOfTokensConsumedInStatement =
                /* The keyword token(s) */ (includesCallKeyword ? 2 : 1) +
                /* The individual target expressions tokens */ targetDetails.Sum(tokenSet => tokenSet.Item2) +
                /* Any argument separaters between target expressions */ (targetExpressionsTokenSets.Count - 1) +
                /* Any extra brackets we removed due to CALL being present */ (includesCallKeyword ? 2 : 0);
            tokens.RemoveRange(0, numberOfTokensConsumedInStatement);
            if (tokens.Any())
            {
                // When getEntryList was first called, it may have returned content when it hit an end-of-statement or when there were no more tokens to consume
                // (so long as the tokens it DID consume were valid and no brackets were mismatched - otherwise it would have thrown an exception). If it was the
                // end of the content then there will be nothing to remove after the tokens occupied by this statement. If it was an end-of-statement token then
                // we need to remove that too since we've effectively processed it.
                tokens.RemoveAt(0);
            }
            return new EraseStatement(
                targetDetails.Select(t => t.Item1),
                keywordLineIndex
            );
        }

        /// <summary>
        /// Take a set of tokens that represent a single ERASE target and extract some information from it - whether it's wrapped in brackets (one of the many
        /// runtime-error-raising failure cases), what the primary "target" is and any arguments. Amongst the other failure cases are property access (eg.
        /// ERASE a.Name) and non-array-index calls (eg. a(0) if a is not an array). These failure cases will bw checked for and expressed by the translation
        /// process, for now we just to represent what is present.
        /// </summary>
        private Tuple<EraseStatement.TargetDetails, int> GetTargetExpressionDetailsWithNumberOfTokensConsumed(IEnumerable<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");

            var tokensArray = tokens.ToArray();
            if (!tokensArray.Any())
                throw new ArgumentException("zero tokens - invalid");
            if (tokensArray.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in tokens set");
            if (tokensArray.Any(t => (!(t is AtomToken)) && (!(t is DateLiteralToken)) && (!(t is StringToken))))
            {
                // The StringBreaker has reduced the original content to AtomToken, DateLiteralToken, StringToken values, which may be parsed here, but it's also
                // generated comment tokens and end-of-statement tokens and they must not appear here in order for it to be valid content (the token set passed
                // in here should be the tokens to the end of this statement only)
                throw new ArgumentException("Unexpected token type - only expect AtomToken, DateLiteralToken or StringToken at this point in the process");
            }

            // It's not valid for targets to be wrapped in brackets, but we still want to analyse what's inside. So strip any off, but record that this was done
            // so that it can be taken into account later in the translation process.
            var bracesRemoved = new List<IToken>();
            while (tokensArray.First() is OpenBrace)
            {
                if (!(tokens.Last() is CloseBrace))
                    throw new ArgumentException("Mismatched brackets on ERASE statement on line " + (tokensArray[0].LineIndex + 1));
                bracesRemoved.Add(tokens.First());
                bracesRemoved.Add(tokens.Last());
                tokensArray = tokensArray.Skip(1).Take(tokensArray.Length - 2).ToArray();
                if (!tokensArray.Any())
                    break;
            }
            if (bracesRemoved.Any() && !tokensArray.Any())
            {
                return Tuple.Create(
                    new EraseStatement.TargetDetails(
                        new Expression(new[] { bracesRemoved.First(), bracesRemoved.Last() }),
                        argumentsIfAny: null,
                        wrappedInBraces: true
                    ),
                    bracesRemoved.Count // = numberOfTokensConsumed
                );
            }

            var targetTokens = tokensArray.TakeWhile(token => !(token is OpenBrace)).ToArray();
            var targetArgumentTokens = tokensArray.Skip(targetTokens.Length);
            IEnumerable<Expression> targetArgumentsIfAny;
            int numberOfTokensInArguments;
            if (targetArgumentTokens.Any())
            {
                var openBrace = (OpenBrace)targetArgumentTokens.First();
                var closeBrace = targetArgumentTokens.Last() as CloseBrace;
                if (closeBrace == null)
                    throw new Exception("Invalid token sequence, mismatched brackets on line (" + (openBrace.LineIndex + 1) + ")");

                targetArgumentsIfAny = base.getEntryList(tokensArray, 2, closeBrace)
                    .Select(argumentTokens => new Expression(argumentTokens))
                    .ToArray();
                
                // The number of tokens consumed is 2 (for the braces) + the total number of those in the argument expressions + the separator tokens between arguments
                numberOfTokensInArguments = 2 + targetArgumentsIfAny.Sum(a => a.Tokens.Count()) + (targetArgumentsIfAny.Count() - 1);
            }
            else
            {
                targetArgumentsIfAny = null;
                numberOfTokensInArguments = 0;
            }

            return Tuple.Create(
                new EraseStatement.TargetDetails(
                    new Expression(targetTokens),
                    targetArgumentsIfAny,
                    wrappedInBraces: bracesRemoved.Any()
                ),
                targetTokens.Length + numberOfTokensInArguments + bracesRemoved.Count // = numberOfTokensConsumer
            );
        }
    }
}
