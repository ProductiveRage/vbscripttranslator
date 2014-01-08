using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Compat;
using CSharpWriter.CodeTranslation.Extensions;
using CSharpWriter.CodeTranslation.StatementTranslation;
using CSharpWriter.Lists;
using CSharpWriter.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.BlockTranslators
{
    public class ClassBlockTranslator : CodeBlockTranslator
    {
		public ClassBlockTranslator(
            CSharpName supportRefName,
            CSharpName envClassName,
            CSharpName envRefName,
            CSharpName outerClassName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
			ITranslateIndividualStatements statementTranslator,
			ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
            : base(supportRefName, envClassName, envRefName, outerClassName, outerRefName, nameRewriter, tempNameGenerator, statementTranslator, valueSettingStatementTranslator, logger) { }

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
					scopeAccessInformation.Extend(classBlock, classBlock.Statements.ToNonNullImmutableList(), ScopeLocationOptions.WithinClass),
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
                new TranslatedStatement("[ComVisible(true)]", indentationDepth),
                new TranslatedStatement("[SourceClassName(" + classBlock.Name.Content.ToLiteral() + ")]", indentationDepth),
                new TranslatedStatement("public class " + className, indentationDepth),
                new TranslatedStatement("{", indentationDepth),
                new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionality).Name + " " + _supportRefName.Name + ";", indentationDepth + 1),
                new TranslatedStatement("private readonly " + _envClassName.Name + " " + _envRefName.Name + ";", indentationDepth + 1),
                new TranslatedStatement("private readonly " + _outerClassName.Name + " " + _outerRefName.Name + ";", indentationDepth + 1),
                new TranslatedStatement(
                    string.Format(
                        "public {0}({1} compatLayer, {2} env, {3} outer){4}",
                        className,
                        typeof(IProvideVBScriptCompatFunctionality).Name,
                        _envClassName.Name,
                        _outerClassName.Name,
                        inheritance
                    ),
                    indentationDepth + 1
                ),
                new TranslatedStatement("{", indentationDepth + 1),
                new TranslatedStatement("if (compatLayer == null)", indentationDepth + 2),
                new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", indentationDepth + 3),
                new TranslatedStatement("if (env == null)", indentationDepth + 2),
                new TranslatedStatement("throw new ArgumentNullException(\"env\");", indentationDepth + 3),
                new TranslatedStatement("if (outer == null)", indentationDepth + 2),
                new TranslatedStatement("throw new ArgumentNullException(\"outer\");", indentationDepth + 3),
                new TranslatedStatement("this." + _supportRefName.Name + " = compatLayer;", indentationDepth + 2),
                new TranslatedStatement("this." + _envRefName.Name + " = env;", indentationDepth + 2),
                new TranslatedStatement("this." + _outerRefName.Name + " = outer;", indentationDepth + 2),
                new TranslatedStatement("}", indentationDepth + 1),
                new TranslatedStatement("", indentationDepth + 1)
            };
		}
	}
}
