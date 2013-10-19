using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Handlers;

namespace VBScriptTranslator.LegacyParser.CodeBlocks
{
    public class CodeBlockHandler
    {
        private IEnumerable<string[]> blockEnds;
        public CodeBlockHandler(IEnumerable<string[]> optionalBlockEnds)
        {
            // Null blockEnds value is valid - for the root code block - but if a non-null
            // list IS specified then every sequence must be non-null, have at least one
            // element and each token in the end sequences must be non-null
            if (optionalBlockEnds == null)
            {
                this.blockEnds = null;
                return;
            }

            var blockEnds = new List<string[]>();
            foreach (var endSequence in optionalBlockEnds)
            {
                var endSequenceClone = (endSequence ?? new string[0]).ToArray();
                if (!endSequence.Any())
                    throw new ArgumentException("Invalid BlockEnd sequence specified: null or blank");
                if (endSequenceClone.Any(s => s == null))
                    throw new ArgumentException("Invalid BlockEnd sequence specified: null value within sequence");
                blockEnds.Add(endSequenceClone);
            }
            this.blockEnds = blockEnds;
        }

        /// <summary>
        /// Try to process the token stream using the AbstractBlockHandler classes - this will raise an exception if none of the handlers can process
        /// the token stream at any point. If it processes the entire stream, the output parameter endSequenceMet will be set as null. If the processing
        /// ends when an end sequence (defined at class initialisation) is met, that sequence will be passed out. This will never return null. Note that
        /// the token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated.
        /// </summary>
        public List<ICodeBlock> Process(List<IToken> tokens, out string[] endSequenceMet)
        {
            var handlers = new AbstractBlockHandler[]
            {
				new OptionExplicitHandler(),

                new ClassHandler(),
                new FunctionHandler(),
                
                // This needs to come after FunctionHandler due to the Private/Public keyword overlap
                new DimHandler(),

                new DoHandler(),
                new WhileHandler(),

                new ForHandler(),

                new IfHandler(),
                new SelectHandler(),

                new ExitHandler(),

                new OnErrorHandler(),

                new RandomizeHandler(),

                new NoSpecialCaseHandler(),
                
                // This should always be used as a last resort
                new StatementHandler()
            };

            var codeBlocks = new List<ICodeBlock>();
            while (tokens.Count > 0)
            {
                // Check for sequence-end for current code block
                if (atBlockEnd(tokens, out endSequenceMet))
                    return codeBlocks;

                // If not sequence-end, try to process content
                ICodeBlock codeBlock = null;
                foreach (var handler in handlers)
                {
                    codeBlock = handler.Process(tokens);
                    if (codeBlock != null)
                    {
                        codeBlocks.Add(codeBlock);
                        break;
                    }
                }
                if (codeBlock == null)
                    throw new Exception("No handler able to handle token stream at current position");
            }
            endSequenceMet = null;
            return codeBlocks;
        }

        /// <summary>
        /// Have we encountered a token sequence that marks the end for this code block?
        /// </summary>
        private bool atBlockEnd(List<IToken> tokens, out string[] endSequenceMet)
        {
            if (this.blockEnds != null)
            {
                foreach (string[] endSequence in this.blockEnds)
                {
                    if (tokens.Count >= endSequence.Length)
                    {
                        int matchCount = 0;
                        for (int index = 0; index < endSequence.Length; index++)
                        {
                            if (tokens[index].Content.ToUpper() == endSequence[index].ToUpper())
                                matchCount++;
                        }
                        if (matchCount == endSequence.Length)
                        {
                            endSequenceMet = endSequence;
                            return true;
                        }
                    }
                }
            }
            endSequenceMet = null;
            return false;
        }
    }
}
