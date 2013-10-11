using System;
using System.Collections.Generic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class FunctionBlock : AbstractFunctionBlock
    {
        public FunctionBlock(
            bool isPublic,
            bool isDefault,
            string name,
            List<Parameter> parameters,
            List<ICodeBlock> statements)
            : base(isPublic, isDefault, name, parameters, statements) { }

        protected override string keyWord
        {
            get { return "Function"; }
        }
    }
}
