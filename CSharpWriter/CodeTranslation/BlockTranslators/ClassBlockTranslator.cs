using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions;
using VBScriptTranslator.CSharpWriter.CodeTranslation.StatementTranslation;
using VBScriptTranslator.CSharpWriter.Lists;
using VBScriptTranslator.CSharpWriter.Logging;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.RuntimeSupport.Compat;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.BlockTranslators
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

			// Outside of classes, functions may be declared multiple times - in which case the last one will take precedence and it will be as if the
			// others never existed. Within classes, however, duplicated names are not allowed and would result in a "Name redefined" compile-time error.
			/* TODO: Fix this to allow properties and functions - no name may be present for both and none may have multiple get, let, set accessors
			var classFunctionsWithDuplicatedNames = classBlock.Statements
				.Where(statement => statement is AbstractFunctionBlock)
				.Cast<AbstractFunctionBlock>()
				.GroupBy(function => function.Name.Content, StringComparer.OrdinalIgnoreCase)
				.Where(group => group.Count() > 1)
				.Select(group => group.Key);
			if (classFunctionsWithDuplicatedNames.Any())
				throw new ArgumentException("The following function name are repeated with class " + classBlock.Name.Content + " (which is invalid): " + string.Join(", ", classFunctionsWithDuplicatedNames));
			 */

			// Apply Class_Initialize / Class_Terminate validation - if they appear then they must be SUBs (not FUNCTIONs) and they may not have arguments
			var classInitializeMethodIfAny = classBlock.Statements
				.Where(statement => statement is AbstractFunctionBlock)
				.Cast<AbstractFunctionBlock>()
				.FirstOrDefault(function => function.Name.Content.Equals("Class_Initialize", StringComparison.OrdinalIgnoreCase));
			if (classInitializeMethodIfAny != null)
			{
				if (!(classInitializeMethodIfAny is SubBlock) || classInitializeMethodIfAny.Parameters.Any())
					throw new ArgumentException("A class' Class_Initialize may only be a Sub (not a Function) with zero arguments - this is not the case in class " + classBlock.Name.Content);
				if (!classInitializeMethodIfAny.Statements.Any(s => !(s is INonExecutableCodeBlock)))
				{
					// If Class_Initialize doesn't do anything then there's no point adding the complexity of calling it (same for Class_Terminate,
					// except that there's even more complexity that can be avoided in that case - see below)
					classInitializeMethodIfAny = null;
				}
			}
			var classTerminateMethodIfAny = classBlock.Statements
				.Where(statement => statement is AbstractFunctionBlock)
				.Cast<AbstractFunctionBlock>()
				.FirstOrDefault(function => function.Name.Content.Equals("Class_Terminate", StringComparison.OrdinalIgnoreCase));
			if (classTerminateMethodIfAny != null)
			{
				if (!(classTerminateMethodIfAny is SubBlock) || classTerminateMethodIfAny.Parameters.Any())
					throw new ArgumentException("A class' Class_Terminate may only be a Sub (not a Function) with zero arguments - this is not the case in class " + classBlock.Name.Content);
				if (!classTerminateMethodIfAny.Statements.Any(s => !(s is INonExecutableCodeBlock)))
				{
					// If Class_Terminate doesn't do anything then there's no point adding the complexity involved in supporting it
					classTerminateMethodIfAny = null;
				}
			}

			// Any error-trapping in the parent scope will not apply within the class and will have to be set explicitly within the methods and
			// properties if required
			scopeAccessInformation = scopeAccessInformation.SetErrorRegistrationToken(null);

			var classContentTranslationResult = TranslateForClass(
				classBlock.Statements.ToNonNullImmutableList(),
				scopeAccessInformation.Extend(classBlock, classBlock.Statements.ToNonNullImmutableList()),
				indentationDepth + 1
			);
			if (classContentTranslationResult.UndeclaredVariablesAccessed.Any())
			{
				// Valid VBScript would not allow an undeclared variables to be present at this point. Any undeclared variables would be become
				// implicitly declared within functions or properties and may not exist at all outside of them (the only things allows within a
				// class are explicit variable declarations - Dim / Private / Public - and functions / properties).
				throw new ArgumentException("Invalid content - it should not be possible for there to be any undeclared variables within a class that aren't within one of its functions or properties");
			}
			var explicitVariableDeclarationsFromWithClass = classContentTranslationResult.ExplicitVariableDeclarations;
			classContentTranslationResult = new TranslationResult(
				classContentTranslationResult.TranslatedStatements,
				new NonNullImmutableList<VariableDeclaration>(), // The ExplicitVariableDeclarations will be translated separately below
				new NonNullImmutableList<NameToken>() // We've just confirmed that there will be no UndeclaredVariablesAccessed references
			);
			var translatedClassHeaderContent = TranslateClassHeader(
				classBlock,
				scopeAccessInformation,
				explicitVariableDeclarationsFromWithClass,
				(classInitializeMethodIfAny == null) ? null : classInitializeMethodIfAny.Name,
				(classTerminateMethodIfAny == null) ? null : classTerminateMethodIfAny.Name,
				indentationDepth
			);
			if (classContentTranslationResult.TranslatedStatements.Any())
				translatedClassHeaderContent = translatedClassHeaderContent.Concat(new[] { new TranslatedStatement("", 0, classBlock.Name.LineIndex) });
			var lineIndexOfLastStatement = classContentTranslationResult.TranslatedStatements.Any()
				? classContentTranslationResult.TranslatedStatements.Max(s => s.LineIndexOfStatementStartInSource)
				: classBlock.Name.LineIndex;
			return TranslationResult.Empty
				.Add(translatedClassHeaderContent)
				.Add(classContentTranslationResult)
				.Add(new TranslatedStatement("}", indentationDepth, lineIndexOfLastStatement));
		}

		private TranslationResult TranslateForClass(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
		{
			if (blocks == null)
				throw new ArgumentNullException("block");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			// Note: DIM statements are allowed between in the space inside a CLASS but outside of any FUNCTION or PROPERTY within that CLASS, but CONST is not
			return base.TranslateCommon(
				new BlockTranslationAttempter[]
				{
					base.TryToTranslateBlankLine,
					base.TryToTranslateComment,
					base.TryToTranslateDim,
					base.TryToTranslateFunctionPropertyOrSub
				}.ToNonNullImmutableList(),
				blocks,
				scopeAccessInformation,
				indentationDepth
			);
		}

		private IEnumerable<TranslatedStatement> TranslateClassHeader(
			ClassBlock classBlock,
			ScopeAccessInformation scopeAccessInformation,
			NonNullImmutableList<VariableDeclaration> explicitVariableDeclarationsFromWithinClass,
			NameToken classInitializeMethodNameIfAny,
			NameToken classTerminateMethodNameIfAny,
			int indentationDepth)
		{
			if (classBlock == null)
				throw new ArgumentNullException("classBlock");
			if (scopeAccessInformation == null)
				throw new ArgumentNullException("scopeAccessInformation");
			if (explicitVariableDeclarationsFromWithinClass == null)
				throw new ArgumentNullException("explicitVariableDeclarationsFromWithClass");
			if (indentationDepth < 0)
				throw new ArgumentOutOfRangeException("indentationDepth", "must be zero or greater");

			// C# doesn't support nameed indexed properties, so if there are any Get properties with a parameter or any Let/Set properties
			// with multiple parameters (they need at least; the value to set) then we'll have to get creative - see notes in the
			// TranslatedPropertyIReflectImplementation class
			string inheritanceChainIfAny;
			if (classBlock.Statements.Where(s => s is PropertyBlock).Cast<PropertyBlock>().Any(p => p.IsPublic && p.IsIndexedProperty()))
				inheritanceChainIfAny = " : " + typeof(TranslatedPropertyIReflectImplementation).Name; // Assume that the namespace is available for this attribute
			else
				inheritanceChainIfAny = "";

			var className = _nameRewriter.GetMemberAccessTokenName(classBlock.Name);
			IEnumerable<TranslatedStatement> classInitializeCallStatements;
			if (classInitializeMethodNameIfAny == null)
				classInitializeCallStatements = new TranslatedStatement[0];
			else
			{
				// When Class_Initialize is called, it is treated as if ON ERROR RESUME NEXT is applied around the call - the first error will
				// result in the method exiting, but the error won't be propagated up (so the caller will continue as if it hadn't happened).
				// Wrapping the call in a try..catch simulates this behaviour (there is similar required for Class_Terminate).
				classInitializeCallStatements = new[] {
					new TranslatedStatement(
						"try { " + _nameRewriter.GetMemberAccessTokenName(classInitializeMethodNameIfAny) + "(); }",
						indentationDepth + 2,
						classBlock.Name.LineIndex
					),
					new TranslatedStatement("catch(Exception e)", indentationDepth + 2, classBlock.Name.LineIndex),
					new TranslatedStatement("{", indentationDepth + 2, classBlock.Name.LineIndex),
					new TranslatedStatement(_supportRefName.Name + ".SETERROR(e);", indentationDepth + 3, classBlock.Name.LineIndex),
					new TranslatedStatement("}", indentationDepth + 2, classBlock.Name.LineIndex)
				};
			}
			TranslatedStatement[] disposeImplementationStatements;
			CSharpName disposedFlagNameIfAny;
			if (classTerminateMethodNameIfAny == null)
			{
				disposeImplementationStatements = new TranslatedStatement[0];
				disposedFlagNameIfAny = null;
			}
			else
			{
				// If this class has a Clas_Terminate method, then the closest facsimile is a finalizer. However, in C# the garbage collector means
				// that the execution of the finalizer is non-deterministic, while the VBScript interpreter's garbage collector resulted in these
				// being executed predictably. In case a translated class requires this behaviour be maintained somehow, the best thing to do is
				// to make the class implement IDisposable. Translated calling code will not know when to call Dispose on the classes written
				// here, but if they are called from any other C# code written directly against the translated output, having the option of
				// taking advantage of the IDisposable interface may be useful.
				// - Note that when the finalizer is executed, the call to Dispose (which then calls the Class_Terminate method) is wrapped in
				//   a try..catch for the same reason as the Class_Initialize call, as explained above. The only difference is that here, when
				//   we call _.SETERROR(e) there is - technically - a chance that the "_" reference will have been finalised. Managed references
				//   should not be accessed in the finaliser if we're following the rules to the letter, but this entire structure around trying
				//   to emulate Class_Terminate with a finaliser is a series of compromises, so this is the best we can do. The SETERROR call is
				//   ALSO wrapped in a try..catch, just in case that reference really is no longer available.
				// - Also note that IDisposable's public Dispose() method is explicitly implemented and that the method that this and the
				//   finaliser call is named by "_tempNameGenerator" reference and that it is "private" instead of the more common (for
				//   correct implementations of the disposable pattern) "protected virtual". This is explained below, where the class
				//   header is generated.
				disposedFlagNameIfAny = _tempNameGenerator(new CSharpName("_disposed"), scopeAccessInformation);
				var disposeMethodName = _tempNameGenerator(new CSharpName("Dispose"), scopeAccessInformation);
				disposeImplementationStatements = new[]
				{
					new TranslatedStatement("~" + className + "()", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("{", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("try { " + disposeMethodName.Name + "(false); }", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("catch(Exception e)", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("{", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("try { " + _supportRefName.Name + ".SETERROR(e); } catch { }", indentationDepth + 3, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("}", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("}", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("void IDisposable.Dispose()", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("{", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement(disposeMethodName.Name + "(true);", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("GC.SuppressFinalize(this);", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("}", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("private void " + disposeMethodName.Name + "(bool disposing)", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("{", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("if (" + disposedFlagNameIfAny.Name + ")", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("return;", indentationDepth + 3, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("if (disposing)", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement(_nameRewriter.GetMemberAccessTokenName(classTerminateMethodNameIfAny) + "();", indentationDepth + 3, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement(disposedFlagNameIfAny.Name + " = true;", indentationDepth + 2, classTerminateMethodNameIfAny.LineIndex),
					new TranslatedStatement("}", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex)
				};
				if (inheritanceChainIfAny == "")
					inheritanceChainIfAny = " : ";
				else
					inheritanceChainIfAny += ", ";
				inheritanceChainIfAny += "IDisposable";
			}

			// The class is sealed to make the IDisposable implementation easier (where appropriate) - the recommended way to implement IDisposable
			// is for there to be a "protected virtual void Dispose(bool disposing)" method, for use by the IDisposable's public Dispose() method,
			// by the finaliser and by any derived types. However, there may be a "Dispose" method that comes from the VBScript source. We can work
			// around this by explicitly implementating "IDisposable.Dispose" and by using the "_tempNameGenerator" reference to get a safe-to-use
			// method name, for the boolean argument "Dispose" method - but then it won't follow the recommended pattern and be a method name
			// "Dispose" that derived types can use. The easiest way around that is to make the classes sealed and then there are no derived
			// types to worry about (this could be seen to be a limitation on the translated code, but since it's all being translated from
			// VBScript where all types are dynamic, one class could just be swapped out for another entirely different one as long as it
			// has the same methods and properties on it).
			var classHeaderStatements = new List<TranslatedStatement>
			{
				new TranslatedStatement("[ComVisible(true)]", indentationDepth, classBlock.Name.LineIndex),
				new TranslatedStatement("[SourceClassName(" + classBlock.Name.Content.ToLiteral() + ")]", indentationDepth, classBlock.Name.LineIndex),
				new TranslatedStatement("public sealed class " + className + inheritanceChainIfAny, indentationDepth, classBlock.Name.LineIndex),
				new TranslatedStatement("{", indentationDepth, classBlock.Name.LineIndex),
				new TranslatedStatement("private readonly " + typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name + " " + _supportRefName.Name + ";", indentationDepth + 1, classBlock.Name.LineIndex),
				new TranslatedStatement("private readonly " + _envClassName.Name + " " + _envRefName.Name + ";", indentationDepth + 1, classBlock.Name.LineIndex),
				new TranslatedStatement("private readonly " + _outerClassName.Name + " " + _outerRefName.Name + ";", indentationDepth + 1, classBlock.Name.LineIndex)
			};
			if (disposedFlagNameIfAny != null)
			{
				classHeaderStatements.Add(
					new TranslatedStatement("private bool " + disposedFlagNameIfAny.Name + ";", indentationDepth + 1, classBlock.Name.LineIndex)
				);
			}
			classHeaderStatements.AddRange(new[] {
				new TranslatedStatement(
					string.Format(
						"public {0}({1} compatLayer, {2} env, {3} outer)",
						className,
						typeof(IProvideVBScriptCompatFunctionalityToIndividualRequests).Name,
						_envClassName.Name,
						_outerClassName.Name
					),
					indentationDepth + 1,
					classBlock.Name.LineIndex
				),
				new TranslatedStatement("{", indentationDepth + 1, classBlock.Name.LineIndex),
				new TranslatedStatement("if (compatLayer == null)", indentationDepth + 2, classBlock.Name.LineIndex),
				new TranslatedStatement("throw new ArgumentNullException(\"compatLayer\");", indentationDepth + 3, classBlock.Name.LineIndex),
				new TranslatedStatement("if (env == null)", indentationDepth + 2, classBlock.Name.LineIndex),
				new TranslatedStatement("throw new ArgumentNullException(\"env\");", indentationDepth + 3, classBlock.Name.LineIndex),
				new TranslatedStatement("if (outer == null)", indentationDepth + 2, classBlock.Name.LineIndex),
				new TranslatedStatement("throw new ArgumentNullException(\"outer\");", indentationDepth + 3, classBlock.Name.LineIndex),
				new TranslatedStatement(_supportRefName.Name + " = compatLayer;", indentationDepth + 2, classBlock.Name.LineIndex),
				new TranslatedStatement(_envRefName.Name + " = env;", indentationDepth + 2, classBlock.Name.LineIndex),
				new TranslatedStatement(_outerRefName.Name + " = outer;", indentationDepth + 2, classBlock.Name.LineIndex)
			});
			if (disposedFlagNameIfAny != null)
			{
				classHeaderStatements.Add(
					new TranslatedStatement(disposedFlagNameIfAny.Name + " = false;", indentationDepth + 2, classBlock.Name.LineIndex)
				);
			}
			classHeaderStatements.AddRange(
				explicitVariableDeclarationsFromWithinClass.Select(
					v => new TranslatedStatement(
						base.TranslateVariableInitialisation(v, ScopeLocationOptions.WithinClass),
						indentationDepth + 2,
						v.Name.LineIndex
					)
				)
			);
			classHeaderStatements.AddRange(classInitializeCallStatements);
			classHeaderStatements.Add(
				new TranslatedStatement("}", indentationDepth + 1, classBlock.Name.LineIndex)
			);
			if (disposeImplementationStatements.Any())
			{
				classHeaderStatements.Add(new TranslatedStatement("", indentationDepth + 1, classTerminateMethodNameIfAny.LineIndex));
				classHeaderStatements.AddRange(disposeImplementationStatements);
			}
			if (explicitVariableDeclarationsFromWithinClass.Any())
			{
				classHeaderStatements.Add(new TranslatedStatement("", indentationDepth + 1, explicitVariableDeclarationsFromWithinClass.Min(v => v.Name.LineIndex)));
				classHeaderStatements.AddRange(
					explicitVariableDeclarationsFromWithinClass.Select(declaredVariableToInitialise =>
						new TranslatedStatement(
							string.Format(
								"{0} object {1} {{ get; set; }}",
								(declaredVariableToInitialise.Scope == VariableDeclarationScopeOptions.Private) ? "private" : "public",
								_nameRewriter.GetMemberAccessTokenName(declaredVariableToInitialise.Name)
							),
							indentationDepth + 1,
							declaredVariableToInitialise.Name.LineIndex
						)
					)
				);
			}
			return classHeaderStatements;
		}
	}
}
