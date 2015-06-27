using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class ConstStatement : ICodeBlock
    {
        public ConstStatement(IEnumerable<ConstValueInitialisation> values)
        {
            if (values == null)
                throw new ArgumentNullException("values");

            Values = values.ToList().AsReadOnly();
            if (!Values.Any())
                throw new ArgumentException("Empty values set - invalid");
            if (Values.Any(v => v == null))
                throw new ArgumentException("Null reference encountered in values set");
        }

        /// <summary>
        /// This will never be null, empty nor contain any nulls
        /// </summary>
        public IEnumerable<ConstValueInitialisation> Values { get; private set; }

        [Serializable]
        public class ConstValueInitialisation
        {
            public ConstValueInitialisation(NameToken name, IToken value)
            {
                if (name == null)
                    throw new ArgumentNullException("name");
                if (value == null)
                    throw new ArgumentNullException("value");
                
                if (!(value is DateLiteralToken) && !(value is NumericValueToken) && !(value is StringToken))
                {
                    var builtInValueToken = value as BuiltInValueToken;
                    if ((builtInValueToken == null) || !builtInValueToken.IsAcceptableAsConstValue)
                        throw new ArgumentException("Invalid CONST value - must be a literal or supported builtin value");
                }

                Name = name;
                Value = value;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public NameToken Name { get; private set; }

            /// <summary>
            /// This will never be null, it will always be a literal value - a boolean, number, string or date - or one of the acceptable builtin values, such as Empty
            /// or Null (acceptables values have a true IsAcceptableAsConstValue property on BuiltInValueToken instance)
            /// </summary>
            public IToken Value { get; private set; }

            public override string ToString()
            {
                return base.ToString() + ":" + Name;
            }
        }

        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
        /// </summary>
        public virtual string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            return string.Format(
                "{0}Const {1}",
                indenter.Indent,
                string.Join(", ", Values.Select(v => v.Name + " = " + v.Value))
            );
        }
    }
}
