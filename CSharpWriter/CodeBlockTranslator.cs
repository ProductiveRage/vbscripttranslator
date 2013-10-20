using CSharpWriter.Lists;
using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter
{
    public class CodeBlockTranslator
    {
        private readonly CSharpName _supportClassName;
        private readonly VBScriptNameRewriter _nameRewriter;
        private readonly TempValueNameGenerator _tempNameGenerator;
        public CodeBlockTranslator(CSharpName supportClassName, VBScriptNameRewriter nameRewriter, TempValueNameGenerator tempNameGenerator)
        {
            if (supportClassName == null)
                throw new ArgumentNullException("supportClassName");
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (tempNameGenerator == null)
                throw new ArgumentNullException("tempNameGenerator");

            _supportClassName = supportClassName;
            _nameRewriter = nameRewriter;
            _tempNameGenerator = tempNameGenerator;
        }

        /// <summary>
        /// This must be responsible for translating any NameToken into a string that is legal for use as a C# identifier. It not expect a null name reference
        /// and must never return a null value. It is responsible for returning consistent names regardless of the case of the input value, to deal with the
        /// fact that C#  is case-sensitive and VBScript is not.
        /// </summary>
        public delegate CSharpName VBScriptNameRewriter(NameToken name);

        /// <summary>
        /// During translation, temporary variables may be required. This delegate is responsible for returning names that are guaranteed to be unique. The
        /// mechanism for implementing this must work with the VBScriptNameRewriter mechanism since there must be no overlap in the returned values. If an
        /// optionalPrefix value is specified then the returned name must begin with this (if null is specified then it must be ignored).
        /// </summary>
        public delegate CSharpName TempValueNameGenerator(CSharpName optionalPrefix);

        public NonNullImmutableList<TranslatedStatement> Translate(NonNullImmutableList<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException("blocks");

            var translationResult = Translate(
                blocks,
                ScopeAccessInformation.Empty.Extend(ParentConstructTypeOptions.None, blocks),
                0
            );
            translationResult = FlushExplicitVariableDeclarations(translationResult, ParentConstructTypeOptions.None, 0);
            translationResult = FlushUndeclaredVariableDeclarations(translationResult, 0);
            return translationResult.TranslatedStatements;
        }

        private TranslationResult Translate(
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeAccessInformation scopeAccessInformation,
            int indentationDepth)
        {
            if (blocks == null)
                throw new ArgumentNullException("block");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            var translationResult = TranslationResult.Empty;
            for (var index = 0; index < blocks.Count; index++)
            {
                var block = blocks[index];
                var nextBlock = (index == (blocks.Count - 1)) ? null : blocks[index + 1]; // TODO: Is this required??

                if (block is OptionExplicit)
                    continue;

                if (block is BlankLine)
                {
                    translationResult = translationResult.Add(
                        new TranslatedStatement("", indentationDepth)
                    );
                    continue;
                }

                var commentBlock = block as CommentStatement;
                if (commentBlock != null)
                {
                    var translatedCommentContent = "//" + commentBlock.Content;
                    if (block is InlineCommentStatement)
                    {
                        var lastTranslatedStatement = translationResult.TranslatedStatements.LastOrDefault();
                        if ((lastTranslatedStatement != null) && (lastTranslatedStatement.Content != ""))
                        {
                            translationResult = new TranslationResult(
                                translationResult.TranslatedStatements
                                    .RemoveLast()
                                    .Add(new TranslatedStatement(
                                        lastTranslatedStatement.Content + " " + translatedCommentContent,
                                        lastTranslatedStatement.IndentationDepth
                                    )),
                                translationResult.ExplicitVariableDeclarations,
                                translationResult.UndeclaredVariablesAccessed
                            );
                            continue;
                        }
                    }
                    translationResult = translationResult.Add(
                        new TranslatedStatement(translatedCommentContent, indentationDepth)
                    );
                    continue;
                }

                // This covers the DimStatement, ReDimStatement, PrivateVariableStatement and PublicVariableStatement
                var explicitVariableDeclarationBlock = block as DimStatement;
                if (explicitVariableDeclarationBlock != null)
                {
                    translationResult = translationResult.Add(
                        explicitVariableDeclarationBlock.Variables.Select(v => 
                            new VariableDeclaration(
                                v.Name,
                                (explicitVariableDeclarationBlock is PublicVariableStatement)
                                    ? VariableDeclarationScopeOptions.Public
                                    : VariableDeclarationScopeOptions.Private,
                                (v.Dimensions != null) // If the Dimensions set is non-null then this is an array type
                            )
                        )
                    );

                    var areDimensionsRequired = (
                        (explicitVariableDeclarationBlock is ReDimStatement) ||
                        (explicitVariableDeclarationBlock.Variables.Any(v => (v.Dimensions != null) && (v.Dimensions.Count > 0)))
                    );
                    if (!areDimensionsRequired)
                        continue;
                    
                    // TODO: If this is a ReDim then non-constant expressions may be used to set the dimension limits, in which case it may not be moved
                    // (though a default-null declaration SHOULD be added as well as leaving the ReDim translation where it is)
                    // TODO: Need a translated statement if setting dimensions
                    throw new NotImplementedException("Not enabled support for declaring array variables with specifid dimensions yet");
                }

                var classBlock = block as ClassBlock;
                if (classBlock != null)
                {
                    translationResult = translationResult.Add(
                        new TranslatedStatement(TranslateClassHeader(classBlock), indentationDepth),
                        Translate(
                            classBlock.Statements.ToNonNullImmutableList(),
                            scopeAccessInformation.Extend(
                                ParentConstructTypeOptions.Class,
                                classBlock.Statements
                            ),
                            indentationDepth + 1
                        ),
                        new TranslatedStatement("}", indentationDepth)
                    );
                    continue;
                }

                var functionBlock = ((block is FunctionBlock) || (block is SubBlock)) ? block as AbstractFunctionBlock : null;
                if (functionBlock != null)
                {
                    translationResult = translationResult.Add(
                        new TranslatedStatement(TranslateFunctionHeader(functionBlock), indentationDepth),
                        Translate(
                            functionBlock.Statements.ToNonNullImmutableList(),
                            scopeAccessInformation.Extend(
                                ParentConstructTypeOptions.FunctionOrProperty,
                                functionBlock.Statements
                            ),
                            indentationDepth + 1
                        ),
                        new TranslatedStatement("}", indentationDepth)
                    );
                    continue;
                }

                var statementBlock = block as Statement;
                if (statementBlock != null)
                {
                    throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
                }

                var valueSettingStatement = block as ValueSettingStatement;
                if (valueSettingStatement != null)
                {
                    throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");
                }

                throw new NotImplementedException("Not enabled support for " + block.GetType() + " yet");

                // TODO
                // - DoBlock
                // - ExitStatement
                // - Expression / Statement / ValueSettingStatement
                // - ForBlock
                // - ForEachBlock
                // - OnErrorResumeNext / OnErrorGoto0
                // - PropertyBlock
                // - RandomizeStatement (see http://msdn.microsoft.com/en-us/library/e566zd96(v=vs.84).aspx when implementing RND)
                // - SelectBlock

                // Error on particular statements encountered out of context
                // - This is a "RegularProcessor"..? (as opposed to ClassProcessor which allows properties but not many other things)
                // 1. Exit Statement (should always be within another construct - eg. Do, For)
                // 2. Properties
            }

            // TODO: Explain..
            if (scopeAccessInformation.ParentConstructType == ParentConstructTypeOptions.NonScopeAlteringConstruct)
                return translationResult;
            
            return FlushExplicitVariableDeclarations(
                translationResult,
                scopeAccessInformation.ParentConstructType,
                indentationDepth
            );
        }

        private string TranslateStatement(Statement statement)
        {
            if (statement == null)
                throw new ArgumentNullException("statement");

            throw new NotImplementedException("Not enabled support for Statement translation yet"); // TODO
        }

        private string TranslateVariableDeclaration(VariableDeclaration variableDeclaration)
        {
            if (variableDeclaration == null)
                throw new ArgumentNullException("variableDeclaration");

            return string.Format(
                "object {0} = {1}null",
                _nameRewriter(variableDeclaration.Name).Name,
                variableDeclaration.IsArray ? "(object[])" : ""
            );
        }

        private string TranslateClassHeader(ClassBlock classBlock)
        {
            if (classBlock == null)
                throw new ArgumentNullException("classBlock");

            throw new NotImplementedException("Not enabled support for Class header translation yet"); // TODO
        }

        private string TranslateFunctionHeader(AbstractFunctionBlock functionBlock)
        {
            if (functionBlock == null)
                throw new ArgumentNullException("functionBlock");

            throw new NotImplementedException("Not enabled support for Function/Sub header translation yet"); // TODO
        }

        private TranslationResult FlushExplicitVariableDeclarations(
            TranslationResult translationResult,
            ParentConstructTypeOptions parentConstructType,
            int indentationDepth)
        {
            // TODO: Consider trying to insert the content after any comments or blank lines?
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (!Enum.IsDefined(typeof(ParentConstructTypeOptions), parentConstructType))
                throw new ArgumentOutOfRangeException("parentConstructType");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            return new TranslationResult(
                translationResult.ExplicitVariableDeclarations
                    .Select(v =>
                         new TranslatedStatement(TranslateVariableDeclaration(v), indentationDepth)
                    )
                    .ToNonNullImmutableList()
                    .AddRange(translationResult.TranslatedStatements),
                new NonNullImmutableList<VariableDeclaration>(),
                translationResult.UndeclaredVariablesAccessed
            );
        }

        /// <summary>
        /// This should only performed at the outer layer (and so no ParentConstructTypeOptions value is required, it is assumed to be None)
        /// </summary>
        private TranslationResult FlushUndeclaredVariableDeclarations(TranslationResult translationResult, int indentationDepth)
        {
            if (translationResult == null)
                throw new ArgumentNullException("translationResult");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

            return new TranslationResult(
                translationResult.UndeclaredVariablesAccessed
                    .Select(v =>
                         new TranslatedStatement(
                            TranslateVariableDeclaration(
                                // Undeclared variables will be specified as non-array types initially (hence the false
                                // value for the isArray argument if the VariableDeclaration constructor call below)
                                new VariableDeclaration(v, VariableDeclarationScopeOptions.Public, false)
                            ),
                            indentationDepth
                        )
                    )
                    .ToNonNullImmutableList()
                    .AddRange(translationResult.TranslatedStatements),
                translationResult.ExplicitVariableDeclarations,
                new NonNullImmutableList<NameToken>()
            );
        }
    }
}
