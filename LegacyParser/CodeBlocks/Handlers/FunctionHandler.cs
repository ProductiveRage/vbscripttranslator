using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Handlers
{
    public class FunctionHandler : AbstractBlockHandler
    {
        private enum BlockType
        {
            Unknown,
            
            PublicSub,
            PrivateSub,
            
            PublicFunction,
            PublicDefaultFunction,
            PrivateFunction,

            PublicPropertyGet,
            PublicDefaultPropertyGet,
            PrivatePropertyGet,

            PublicPropertySet,
            PrivatePropertySet,

            PublicPropertyLet,
            PrivatePropertyLet
        }

        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// Note: This handles both FUNCTION and SUB blocks, since they are essentially the same
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count == 0)
                return null;

            // Look for start of function declaration
            Dictionary<string[], BlockType> matchPatterns = new Dictionary<string[], BlockType>();
            
            // - Sub
            matchPatterns.Add(new string[] { "SUB" }, BlockType.PublicSub);
            matchPatterns.Add(new string[] { "PUBLIC", "SUB" }, BlockType.PublicSub);
            matchPatterns.Add(new string[] { "PRIVATE", "SUB" }, BlockType.PrivateSub);
            
            // - Function
            matchPatterns.Add(new string[] { "FUNCTION" }, BlockType.PublicFunction);
            matchPatterns.Add(new string[] { "PUBLIC", "FUNCTION" }, BlockType.PublicFunction);
            matchPatterns.Add(new string[] { "PUBLIC", "DEFAULT", "FUNCTION" }, BlockType.PublicDefaultFunction);
            matchPatterns.Add(new string[] { "PRIVATE", "FUNCTION" }, BlockType.PrivateFunction);

            // - Property Get
            matchPatterns.Add(new string[] { "PROPERTY", "GET" }, BlockType.PublicPropertyGet);
            matchPatterns.Add(new string[] { "PUBLIC", "PROPERTY", "GET" }, BlockType.PublicPropertyGet);
            matchPatterns.Add(new string[] { "PUBLIC", "DEFAULT", "PROPERTY", "GET" }, BlockType.PublicDefaultPropertyGet);
            matchPatterns.Add(new string[] { "PRIVATE", "PROPERTY", "GET" }, BlockType.PrivatePropertyGet);

            // - Property Let / Set
            matchPatterns.Add(new string[] { "PROPERTY", "LET" }, BlockType.PublicPropertyLet);
            matchPatterns.Add(new string[] { "PROPERTY", "SET" }, BlockType.PublicPropertySet);
            matchPatterns.Add(new string[] { "PUBLIC", "PROPERTY", "LET" }, BlockType.PublicPropertyLet);
            matchPatterns.Add(new string[] { "PUBLIC", "PROPERTY", "SET" }, BlockType.PublicPropertySet);
            matchPatterns.Add(new string[] { "PRIVATE", "PROPERTY", "LET" }, BlockType.PrivatePropertyLet);
            matchPatterns.Add(new string[] { "PRIVATE", "PROPERTY", "SET" }, BlockType.PrivatePropertySet);

            bool match = false;
            int matchPatternLength = 0;
            BlockType blockType = BlockType.Unknown;
            foreach (string[] matchPattern in matchPatterns.Keys)
            {
                if (base.checkAtomTokenPattern(tokens, matchPattern, false))
                {
                    match = true;
                    matchPatternLength = matchPattern.Length;
                    blockType = matchPatterns[matchPattern];
                    break;
                }
            }
            if (!match)
                return null;

            // Now look for the rest (function name, parameters, etc..)
            // - There must be AT LEAST matchPatternLength + 1 tokens (for function
            //   name; the open/close brackets and parameters are optional)
            if (tokens.Count < matchPatternLength + 1)
                return null;
            
            // - Get IsPublic and funcName (ensure funcName token is AtomToken)
            bool isPublic = (tokens[0].Content.ToUpper() != "PRIVATE");
            if (!(tokens[matchPatternLength] is AtomToken))
                return null;
            string funcName = tokens[matchPatternLength].Content;

            // - Get parameters (if specified, they're optional in VBScript) and
            //   remove the tokens we've accounted for for the function header
            List<FunctionBlock.Parameter> parameters
                = getFuncParametersAndRemoveTokens(tokens, matchPatternLength + 1);

            // Determine what the end sequence will look like..
            List<string[]> endSequences = new List<string[]>();
            switch (blockType)
            {
                case BlockType.PublicSub:
                case BlockType.PrivateSub:
                    endSequences.Add(new string[] { "END", "SUB" });
                    break;
                case BlockType.PublicFunction:
                case BlockType.PublicDefaultFunction:
                case BlockType.PrivateFunction:
                    endSequences.Add(new string[] { "END", "FUNCTION" });
                    break;
                case BlockType.PublicPropertyGet:
                case BlockType.PublicDefaultPropertyGet:
                case BlockType.PrivatePropertyGet:
                case BlockType.PublicPropertySet:
                case BlockType.PrivatePropertySet:
                    endSequences.Add(new string[] { "END", "PROPERTY" });
                    break;
                default:
                    throw new Exception("Ended up with invalid BlockType [" + blockType.ToString() + "] - how did this happen?? :S");
            }

            // Get function content
            string[] endSequenceMet;
            CodeBlockHandler codeBlockHandler = new CodeBlockHandler(endSequences);
            List<ICodeBlock> blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new Exception("Didn't find end sequence!");
            
            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if (tokens.Count > 0)
            {
                if (!(tokens[0] is AbstractEndOfStatementToken))
                    throw new Exception("EndOfStatementToken missing after END SUB/FUNCTION");
                else
                    tokens.RemoveAt(0);
            }

            // Return code block instance
            bool isDefault
                = ((blockType == BlockType.PublicDefaultFunction)
                || (blockType == BlockType.PublicDefaultPropertyGet));

            if ((blockType == BlockType.PublicSub)
            || (blockType == BlockType.PrivateSub))
                return new SubBlock(isPublic, isDefault, funcName, parameters, blockContent);

            else if ((blockType == BlockType.PublicFunction)
            || (blockType == BlockType.PublicDefaultFunction)
            || (blockType == BlockType.PrivateFunction))
                return new FunctionBlock(isPublic, isDefault, funcName, parameters, blockContent);

            else if ((blockType == BlockType.PublicPropertyGet)
            || (blockType == BlockType.PublicDefaultPropertyGet)
            || (blockType == BlockType.PrivatePropertyGet))
                return new PropertyBlock(isPublic, true, funcName, PropertyBlock.PropertyType.Get, parameters, blockContent);

            else if ((blockType == BlockType.PublicPropertySet)
            || (blockType == BlockType.PrivatePropertySet))
                return new PropertyBlock(isPublic, true, funcName, PropertyBlock.PropertyType.Set, parameters, blockContent);

            if ((blockType == BlockType.PublicPropertyLet)
            || (blockType == BlockType.PrivatePropertyLet))
                return new PropertyBlock(isPublic, true, funcName, PropertyBlock.PropertyType.Let, parameters, blockContent);

            else
                throw new Exception("Unrecognised BlockType [" + blockType.ToString() + "] - how did this happen??");
        }

        private List<FunctionBlock.Parameter> getFuncParametersAndRemoveTokens(
            List<IToken> tokens,
            int offset)
        {
            // Note: Could probably do this more easily with base.getEntryList, but wrote
            // this method before that one!

            bool byRef;
            string name;
            bool isArray;
            IToken token;
            List<FunctionBlock.Parameter> parameters
                = new List<FunctionBlock.Parameter>();

            // Check whether we have parameters defined (the entire parameters section
            // if optional - including the brackets)
            if (!base.isEndOfStatement(tokens, offset))
            {
                token = base.getToken_AtomOnly(tokens, offset);
                if (token.Content == "(")
                {
                    // Look for parameter content
                    offset++;
                    while (true)
                    {
                        // Have we reached the closing bracket?
                        token = base.getToken_AtomOnly(tokens, offset);
                        if (token.Content == ")")
                        {
                            // Yes - exit loop
                            offset++;
                            break;
                        }

                        // No - try to extract parameter data (this will throw an
                        // exception if the content is invalid)
                        List<IToken> paramTokens = getParamTokens(tokens, offset, out byRef, out name, out isArray);
                        if ((paramTokens == null) || (paramTokens.Count == 0))
                            throw new Exception("Unexpected content from getParamsToken");
                        parameters.Add(new FunctionBlock.Parameter(byRef, name, isArray));
                        offset += paramTokens.Count;

                        // Next token should be close bracket (handled above) or
                        // parameter separator (check for now)
                        token = base.getToken_AtomOnly(tokens, offset);
                        if (token.Content == ",")
                            offset++;
                    }
                }
            }

            // Ensure next token is EndOfStatement
            if (!base.isEndOfStatement(tokens, offset))
                throw new Exception("Expected EndOfStatementToken after function declaration");
            
            // Trim off handled tokens and return parameter data
            offset++;
            tokens.RemoveRange(0, offset);
            return parameters;
        }

        /// <summary>
        /// Try to determine parameter form - may have "ByRef" or "ByVal" prefix, may have
        /// "()" array declaration. If we've got here then we've determined that there should
        /// be a parameter to grab data for. Raise an exception for invalid content. Return
        /// the tokens that are required to define the detected paramter.
        /// </summary>
        private List<IToken> getParamTokens(List<IToken> tokens, int offset, out bool byRef, out string name, out bool isArray)
        {
            // We need to return the tokens related to the current parameter so that the
            // caller knows how many tokens that have been processed
            List<IToken> paramTokens = new List<IToken>();
            name = null;

            // Determine ByRef / ByVal (default to ByVal if not specified)
            IToken token = base.getToken_AtomOnly(tokens, offset);
            if (token.Content.ToUpper() == "BYREF")
            {
                byRef = true;
                paramTokens.Add(token);
                offset++;
            }
            else if (token.Content.ToUpper() == "BYVAL")
            {
                byRef = false;
                paramTokens.Add(token);
                offset++;
            }
            else
                byRef = true; // VBScript defaults to ByRef behaviour

            // Grab parameter name
            token = base.getToken_AtomOnly(tokens, offset);
            name = token.Content;
            paramTokens.Add(token);
            offset++;

            // Check for open-close brackets (ie. whether parameter is array or not)
            token = base.getToken_AtomOnly(tokens, offset);
            if (token.Content == "(")
            {
                paramTokens.Add(token);
                offset++;
                token = base.getToken_AtomOnly(tokens, offset);
                if (token.Content != ")")
                    throw new Exception("Invalid content in function array parameter declaration");
                paramTokens.Add(token);
                offset++;
                isArray = true;
            }
            else
                isArray = false;

            return paramTokens;
        }
    }
}
