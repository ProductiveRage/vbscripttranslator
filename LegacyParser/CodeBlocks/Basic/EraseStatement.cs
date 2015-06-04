using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    /// <summary>
    /// TODO: Explain..
    /// </summary>
    [Serializable]
    public class EraseStatement : IHaveNonNestedExpressions
    {
        public EraseStatement(IEnumerable<TargetDetails> targets, int keywordLineIndex)
        {
            if (targets == null)
                throw new ArgumentNullException("targets");
            if (keywordLineIndex < 0)
                throw new ArgumentOutOfRangeException("keywordLineIndex");

            Targets = targets.ToList().AsReadOnly();
            if (Targets.Any(t => t == null))
                throw new ArgumentException("Encountered null reference in targets set");
            KeywordLineIndex = keywordLineIndex;
        }

        /// <summary>
        /// This will never be null nor contain any null references. VBScript's Erase function only supports a single target, which must be an array type - however
        /// if there are zero or multiple targets specified then it will fail at runtime, not at compile time, so it is a statement that must be describable, even
        /// if we know it will fail at runtime (unless it's wrapped in an ON ERROR RESUME NEXT).
        /// </summary>
        public IEnumerable<TargetDetails> Targets { get; private set; }

        /// <summary>
        /// This will be useful for the runtime error message that must be generated if there are zero target
        /// </summary>
        public int KeywordLineIndex { get; private set; }

        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> IHaveNonNestedExpressions.NonNestedExpressions
        {
            get
            {
                foreach (var target in Targets)
                {
                    yield return target.Target;
                    if (target.ArgumentsIfAny != null)
                    {
                        foreach (var argument in target.ArgumentsIfAny)
                            yield return argument;
                    }
                }
            }
        }

        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            var output = new StringBuilder();
            output.Append(indenter.Indent);
            output.Append("ERASE ");
            foreach (var indexedTarget in Targets.Select((t, i) => new { Index = i, Value = t }))
            {
                if (indexedTarget.Index > 0)
                    output.Append(", ");
                
                if (indexedTarget.Value.WrappedInBraces)
                    output.Append("(");

                output.Append(indexedTarget.Value.Target.GenerateBaseSource(NullIndenter.Instance));
                if (indexedTarget.Value.ArgumentsIfAny != null)
                {
                    output.Append("(");
                    foreach (var indexedArgument in indexedTarget.Value.ArgumentsIfAny.Select((a, i) => new { Index = i, Value = a }))
                    {
                        if (indexedArgument.Index > 0)
                            output.Append(", ");
                        output.Append(indexedArgument.Value.GenerateBaseSource(NullIndenter.Instance));
                    }
                    output.Append(")");
                }

                if (indexedTarget.Value.WrappedInBraces)
                    output.Append(")");
            }
            return output.ToString();
        }

        [Serializable]
        public class TargetDetails
        {
            public TargetDetails(Expression target, IEnumerable<Expression> argumentsIfAny, bool wrappedInBraces)
            {
                if (target == null)
                    throw new ArgumentNullException("target");

                Target = target;
                ArgumentsIfAny = (argumentsIfAny == null) ? null : argumentsIfAny.ToList().AsReadOnly();
                if ((ArgumentsIfAny != null) && ArgumentsIfAny.Any(a => a == null))
                    throw new ArgumentException("Null reference encountered in targetArgumentsIfAny set");
                WrappedInBraces = wrappedInBraces;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public Expression Target { get; private set; }

            /// <summary>
            /// This is optional and may be null (indicating no arguments and no argument-wrapping brackets) or empty (meaning no arguments but WITH brackets)
            /// or non-empty (in which case there were also brackets around the arguments). If non-null, it will never contain any null references.
            /// </summary>
            public IEnumerable<Expression> ArgumentsIfAny { get; private set; }

            /// <summary>
            /// It's invalid for targets to be wrapped in braces (it will result in a runtime error), so this is important information
            /// </summary>
            public bool WrappedInBraces { get; private set; }
        }
    }
}