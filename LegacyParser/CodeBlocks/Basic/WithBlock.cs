using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class WithBlock : IHaveNestedContent
    {
        public WithBlock(Expression target, IEnumerable<ICodeBlock> content)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (content == null)
                throw new ArgumentNullException("content");

            Target = target;
            Content = content.ToArray();
            if (Content.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in content set");
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public Expression Target { get; private set; }
        
        /// <summary>
        /// This will never be null nor contain any null references
        /// </summary>
        public IEnumerable<ICodeBlock> Content { get; private set; }

        /// <summary>
        /// This is a flattened list of all executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get
            {
                yield return Target;
                foreach (var statement in Content)
                    yield return statement;
            }
        }

        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            var output = new StringBuilder();

            output.Append(indenter.Indent);
            output.Append("WITH ");
            output.AppendLine(Target.GenerateBaseSource(new NullIndenter()));

            foreach (var statement in Content)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            output.Append(indenter.Indent);
            output.Append("END WITH");
            return output.ToString();
        }
    }
}
