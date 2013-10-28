using System;
using System.Collections.Generic;
using System.Linq;
using CSharpSupport;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class ClassBlockTranslator : CodeBlockTranslator
    {
		public ClassBlockTranslator(
            CSharpName supportClassName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator) : base(supportClassName, nameRewriter, tempNameGenerator, statementTranslator) { }

		public TranslationResult Translate(ClassBlock classBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (classBlock == null)
				throw new ArgumentNullException("classBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // TODO: Require a variation of OuterScopeBlockTranslator's RemoveDuplicateFunctions (see notes on that method)
			var translationResult = TranslationResult.Empty.Add(
				TranslateClassHeader(classBlock, indentationDepth)
			);
			translationResult = translationResult.Add(
				Translate(
					classBlock.Statements.ToNonNullImmutableList(),
					scopeAccessInformation.Extend(classBlock, classBlock.Statements),
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
					base.TryToTranslateComment,
					base.TryToTranslateDim,
					base.TryToTranslateFunction,
					base.TryToTranslateProperty
				}.ToNonNullImmutableList(),
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
		}

		private IEnumerable<TranslatedStatement> TranslateClassHeader(ClassBlock classBlock, int indentationDepth)
		{
			if (classBlock == null)
				throw new ArgumentNullException("classBlock");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            // C# doesn't support nameed indexed properties, so if there are any Get properties with a parameter or any Let/Set properties
            // with multiple parameters (they need at least; the value to set) then we'll have to get creative
            string inheritance;
            if (classBlock.Statements.Where(s => s is PropertyBlock).Cast<PropertyBlock>().Any(p => p.IsPublic && p.IsIndexedProperty()))
                inheritance = " : " + typeof(TranslatedPropertyIReflectImplementation).FullName;
            else
                inheritance = "";

			var className = _nameRewriter.GetMemberAccessTokenName(classBlock.Name);
			return new[]
            {
                new TranslatedStatement("[System.Runtime.InteropServices.ComVisible(true)]", indentationDepth),
                new TranslatedStatement("[" + typeof(SourceClassName).FullName + "(" + classBlock.Name.Content.ToLiteral() + ")]", indentationDepth),
                new TranslatedStatement("public class " + className, indentationDepth),
                new TranslatedStatement("{", indentationDepth),
                new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionality).FullName + " " + _supportClassName.Name + ";", indentationDepth + 1),
                new TranslatedStatement("public " + className + "(" + typeof(IProvideVBScriptCompatFunctionality).FullName + " compatLayer)" + inheritance, indentationDepth + 1),
                new TranslatedStatement("{", indentationDepth + 1),
                new TranslatedStatement("if (compatLayer == null)", indentationDepth + 2),
                new TranslatedStatement("throw new ArgumentNullException(compatLayer)", indentationDepth + 3),
                new TranslatedStatement("this." + _supportClassName.Name + " = compatLayer;", indentationDepth + 2),
                new TranslatedStatement("}", indentationDepth + 1),
                new TranslatedStatement("", indentationDepth + 1)
            };
		}
	}
}
