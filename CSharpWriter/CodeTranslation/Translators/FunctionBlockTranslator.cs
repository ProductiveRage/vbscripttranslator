using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSharpSupport;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks;

namespace CSharpWriter.CodeTranslation
{
    public class FunctionBlockTranslator : CodeBlockTranslator
    {
        public FunctionBlockTranslator(
            CSharpName supportClassName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator) : base(supportClassName, nameRewriter, tempNameGenerator, statementTranslator) { }

		public TranslationResult Translate(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (functionBlock == null)
				throw new ArgumentNullException("functionBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			var translationResult = TranslationResult.Empty.Add(
				TranslateFunctionHeader(
					functionBlock,
					(functionBlock is FunctionBlock), // hasReturnValue (true for FunctionBlock, false for SubBlock)
					indentationDepth
				)
			);
			translationResult = translationResult.Add(
				Translate(
					functionBlock.Statements.ToNonNullImmutableList(),
					scopeAccessInformation.Extend(
						ParentConstructTypeOptions.FunctionOrProperty,
						functionBlock.Statements
					),
					indentationDepth + 1
				)
			);
			return translationResult.Add(
				new TranslatedStatement("}", indentationDepth)
			);
		}

		private TranslationResult Translate(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			// TODO
			// - 

			return base.TranslateCommon(blocks, scopeAccessInformation, indentationDepth);
		}

		private IEnumerable<TranslatedStatement> TranslateFunctionHeader(AbstractFunctionBlock functionBlock, bool hasReturnValue, int indentationDepth)
		{
			if (functionBlock == null)
				throw new ArgumentNullException("functionBlock");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			var content = new StringBuilder();
			content.Append(functionBlock.IsPublic ? "public" : "private");
			content.Append(" ");
			content.Append(hasReturnValue ? "object" : "void");
			content.Append(" ");
			content.Append(_nameRewriter.GetMemberAccessTokenName(functionBlock.Name));
			content.Append("(");
			var numberOfParameters = functionBlock.Parameters.Count();
			for (var index = 0; index < numberOfParameters; index++)
			{
				var parameter = functionBlock.Parameters.ElementAt(index);
				if (parameter.ByRef)
					content.Append("ref ");
				content.Append(_nameRewriter.GetMemberAccessTokenName(parameter.Name));
				if (index < (numberOfParameters - 1))
					content.Append(", ");
			}
			content.Append(")");

			var translatedStatements = new List<TranslatedStatement>();
			if (functionBlock.IsDefault)
				translatedStatements.Add(new TranslatedStatement("[" + typeof(IsDefault).FullName + "]", indentationDepth));
			translatedStatements.Add(new TranslatedStatement(content.ToString(), indentationDepth));
			translatedStatements.Add(new TranslatedStatement("{", indentationDepth));
			return translatedStatements;
		}
    }
}
