using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    /// <summary>
    /// This handles both Dim and ReDim statements
    /// </summary>
    public class DimHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            // Note: Dim statement in VBScript can only have constant dimensions specified,
            // variables can not be used (only in ReDim statements). Strings are technically
            // valid (they will be cast to integers, as will decimals) and so is the value
            // -1 (which appears to uninitialise the array), but these are NOT going to be
            // allowed here. Similarly, constant statements are accetable (eg. 1+1), but
            // we're not going to allow them either.
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            // Can we handle the content?
            DimType dimType;
            int tokensConsumed;
            if (!canBeHandled(tokens, out tokensConsumed, out dimType))
                return null;
            if (tokens.Count == tokensConsumed)
                throw new ArgumentException("DimHandler recognises opening AtomToken(s), but there is no following content");

            // Get raw variable content (one token if just variable, three tokens if array
            // without dimensions, four tokens if 1D array, etc..)
            List<List<IToken>> variablesData = base.getEntryList(tokens, tokensConsumed, new EndOfStatementNewLineToken());
            
            // Trim out the opening keyword(s)
            tokens.RemoveRange(0, tokensConsumed);

            // Translate DimStatement.DimVariable objects
            List<DimStatement.DimVariable> variables = new List<DimStatement.DimVariable>();
            foreach (List<IToken> variableData in variablesData)
                variables.Add(translateRawVariableData(variableData));

            // Remove all variable data from token list (and the EndOfStatmentToken)
            foreach (List<IToken> entry in variablesData)
                tokens.RemoveRange(0, entry.Count);
            tokens.RemoveAt(0);

            // Return final object
            if (dimType == DimType.Dim)
                return new DimStatement(variables);
            else if (dimType == DimType.ReDim)
                return new ReDimStatement(false, variables);
            else if (dimType == DimType.ReDimPreserve)
                return new ReDimStatement(true, variables);
            else if (dimType == DimType.Public)
				return new PublicVariableStatement(variables);
			else if (dimType == DimType.Private)
                return new PrivateVariableStatement(variables);
            else
                throw new Exception("Ended up with unexpected DimType value!");
        }

        /// <summary>
        /// Does it look like we can work with the token stream? If so, return some data
        /// regarding what statement type it is, and how many tokens were required to
        /// define that type
        /// </summary>
        private bool canBeHandled(List<IToken> tokens, out int tokensConsumed, out DimType dimType)
        {
            if (base.checkAtomTokenPattern(tokens, new string[] { "DIM" }, false))
            {
                tokensConsumed = 1;
                dimType = DimType.Dim;
                return true;
            }
            if (base.checkAtomTokenPattern(tokens, new string[] { "REDIM" }, false))
            {
                if (base.checkAtomTokenPattern(tokens, new string[] { "REDIM", "PRESERVE" }, false))
                {
                    tokensConsumed = 2;
                    dimType = DimType.ReDimPreserve;
                }
                else
                {
                    tokensConsumed = 1;
                    dimType = DimType.ReDim;
                }
                return true;
            }
            if (base.checkAtomTokenPattern(tokens, new string[] { "PUBLIC" }, false))
            {
                if (!base.checkAtomTokenPattern(tokens, new string[] { "PUBLIC", "FUNCTION" }, false)
                && !base.checkAtomTokenPattern(tokens, new string[] { "PUBLIC", "PROPERTY" }, false)
                && !base.checkAtomTokenPattern(tokens, new string[] { "PUBLIC", "DEFAULT", "PROPERTY" }, false)
                && !base.checkAtomTokenPattern(tokens, new string[] { "PUBLIC", "SUB" }, false))
                {
                    tokensConsumed = 1;
                    dimType = DimType.Public;
                    return true;
                }
            }
            if (base.checkAtomTokenPattern(tokens, new string[] { "PRIVATE" }, false))
            {
                if (!base.checkAtomTokenPattern(tokens, new string[] { "PRIVATE", "FUNCTION" }, false)
                && !base.checkAtomTokenPattern(tokens, new string[] { "PRIVATE", "PROPERTY" }, false)
                && !base.checkAtomTokenPattern(tokens, new string[] { "PRIVATE", "SUB" }, false))
                {
                    tokensConsumed = 1;
                    dimType = DimType.Private;
                    return true;
                }
            }
            tokensConsumed = 0;
            dimType = DimType.Unknown;
            return false;
        }

        private enum DimType
        {
            Unknown,
            Dim,
            ReDim,
            ReDimPreserve,
            Public,
            Private
        }

        private DimStatement.DimVariable translateRawVariableData(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                throw new ArgumentException("zero tokens - invalid");
            foreach (IToken token in tokens)
            {
                if (token == null)
                    throw new Exception("Invalid token - null");
                if ((!(token is AtomToken)) && (!(token is StringToken)))
                    throw new Exception("Invalid token - not AtomToken or StringToken");
            }

            // Get name (if no other content, we're all done!)
            string name = tokens[0].Content;
            if (tokens.Count == 1)
                return new DimStatement.DimVariable(name, null);

            // Ensure next token and last token are "(" and ")"
            if (tokens.Count == 2)
                throw new Exception("Invalid token sequence");
            if ((tokens[1].Content != "(") || (tokens[tokens.Count - 1].Content != ")"))
                throw new Exception("Invalid token sequence");

            // If there were only three tokens, we're all done!
            if (tokens.Count == 3)
                return new DimStatement.DimVariable(name, new List<Expression>());

            // Use base.getEntryList to be flexible and grab dimension declarations
            // as Statement instances
            List<Expression> dimensions = new List<Expression>();
            List<List<IToken>> dimStatements = base.getEntryList(tokens, 2, AtomToken.GetNewToken(")"));
            foreach (List<IToken> dimStatement in dimStatements)
                dimensions.Add(new Expression(dimStatement));

            return new DimStatement.DimVariable(name, dimensions);
        }
    }
}
