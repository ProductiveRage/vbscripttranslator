using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class PropertyBlock : AbstractFunctionBlock
    {
        public PropertyBlock(
            bool isPublic,
            bool isDefault,
            NameToken name,
            PropertyType propType,
            List<Parameter> parameters,
            List<ICodeBlock> statements)
            : base(isPublic, isDefault, name, parameters, statements)
        {
            bool isValid = false;
            foreach (object value in Enum.GetValues(typeof(PropertyType)))
            {
                if (value.Equals(propType))
                {
                    isValid = true;
                    break;
                }
            }
            if (!isValid)
                throw new ArgumentException("Invalid type value specified [" + propType.ToString() + "]");
            PropType = propType;
        }

        public enum PropertyType
        {
            Get,
            Set,
            Let
        }

        public PropertyType PropType { get; private set; }

        protected override string keyWord
        {
            get { return "Property " + PropType.ToString(); }
        }
    }
}
