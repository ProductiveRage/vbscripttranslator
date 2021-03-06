﻿using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class FunctionBlock : AbstractFunctionBlock
    {
        public FunctionBlock(
            bool isPublic,
            bool isDefault,
            NameToken name,
            IEnumerable<Parameter> parameters,
            IEnumerable<ICodeBlock> statements)
            : base(isPublic, isDefault, true, name,parameters, statements) { }

        protected override string keyWord
        {
            get { return "Function"; }
        }
    }
}
