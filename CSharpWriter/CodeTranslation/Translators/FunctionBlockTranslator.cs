using CSharpSupport;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

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
                        functionBlock,
                        _tempNameGenerator(new CSharpName("retVal")),
                        functionBlock.Statements.ToNonNullImmutableList()
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

			return base.TranslateCommon(
				new BlockTranslationAttempter[]
				{
					base.TryToTranslateBlankLine,
					base.TryToTranslateClass,
					base.TryToTranslateComment,
					base.TryToTranslateDim,
					base.TryToTranslateDo,
					base.TryToTranslateExit,
					base.TryToTranslateFor,
					base.TryToTranslateForEach,
					base.TryToTranslateIf,
					base.TryToTranslateRandomize,
					base.TryToTranslateStatementOrExpression,
					base.TryToTranslateSelect,
                    base.TryToTranslateValueSettingStatement
				}.ToNonNullImmutableList(),
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
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
                content.Append("object ");
                content.Append(_nameRewriter.GetMemberAccessTokenName(parameter.Name));
                if (index < (numberOfParameters - 1))
					content.Append(", ");
			}
			content.Append(")");

			var translatedStatements = new List<TranslatedStatement>();
			if (functionBlock.IsDefault)
				translatedStatements.Add(new TranslatedStatement("[" + typeof(IsDefault).FullName + "]", indentationDepth));
            var property = functionBlock as PropertyBlock;
            if ((property != null) && property.IsPublic && property.IsIndexedProperty())
            {
                translatedStatements.Add(
                    new TranslatedStatement(
                        string.Format(
                            "[" + typeof(TranslatedProperty).FullName + "({0})]",
                            property.Name.Content.ToLiteral()
                        ),
                        indentationDepth
                    )
                );
            }
            translatedStatements.Add(new TranslatedStatement(content.ToString(), indentationDepth));
            translatedStatements.Add(new TranslatedStatement("{", indentationDepth));
            return translatedStatements;
		}
    }
}
