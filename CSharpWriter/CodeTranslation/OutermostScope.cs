using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation
{
    /// <summary>
    /// Since the ScopeAccessInformation requires a non-null ScopeDefiningParent, this class is required to wrap statements that are in the outer
    /// most scope in the source content.
    /// </summary>
    public class OutermostScope : IDefineScope
    {
        public OutermostScope(CSharpName wrapperName, NonNullImmutableList<ICodeBlock> codeBlocks)
        {
            if (wrapperName == null)
                throw new ArgumentNullException("wrapperName");
            if (codeBlocks == null)
                throw new ArgumentNullException("codeBlocks");

            Name = new DoNotRenameNameToken(wrapperName.Name, 0);
            AllExecutableBlocks = codeBlocks;
        }

        /// <summary>
        /// This must never be null
        /// </summary>
        public NameToken Name { get; private set; }

        /// <summary>
        /// This must never be null but it may be empty (this may be the names of a a function's arguments, for example)
        /// </summary>
        public IEnumerable<NameToken> ExplicitScopeAdditions { get { return new NameToken[0]; } }

        public ScopeLocationOptions Scope { get { return ScopeLocationOptions.OutermostScope; } }

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
        public IEnumerable<ICodeBlock> AllExecutableBlocks { get; private set; }

        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(ISourceIndentHandler indenter)
        {
            if (indenter == null)
                throw new ArgumentNullException("indenter");

            var writer = new StringBuilder();
            foreach (var block in AllExecutableBlocks)
            {
                if (writer.Length > 0)
                    writer.AppendLine();
                writer.Append(block.GenerateBaseSource(indenter));
            }
            return writer.ToString();
        }
    }
}