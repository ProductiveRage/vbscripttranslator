using System;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class VBScriptNameRewriter_Extensions
    {
        /// <summary>
        /// When trying to access variables, functions, classes, etc.. we need to pass the member's name through the VBScriptNameRewriter. In
        /// most cases this token will be a NameToken which we can pass straight in, but in some cases it may be another type (perhaps a key
        /// word type) and so will have to be wrapped in a NameToken instance before passing through the name rewriter. This extension
		/// method should be used in all places where the VBScriptNameRewriter is used by the CSharpWriter since it allows us to override
		/// its behaviour where required - eg. by using a DoNotRenameNameToken
        /// </summary>
        public static string GetMemberAccessTokenName(this VBScriptNameRewriter nameRewriter, IToken token)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (token == null)
                throw new ArgumentNullException("token");

            // A TargetCurrentClassToken indicates a "Me" (eg. "Me.Name") which can always be translated directly into "this". In VBScript,
            // "Me" is valid even when not explicitly within a VBScript class (it refers to the outermost scope, so "Me.F1()" will try to
            // call a function "F1" in the outermost scope). When the code IS explicitly within a VBScript class, the "Me" reference is
            // the instance of that class. In the translated code, both cases are fine to translate straight into whatever "this" is
            // at runtime.
            if (token is TargetCurrentClassToken)
                return "this";

            var nameToken = (token as NameToken) ?? new ForRenamingNameToken(token.Content, token.LineIndex);
			if (nameToken is DoNotRenameNameToken)
				return nameToken.Content;
            return nameRewriter(nameToken).Name;
        }

        public static bool AreNamesEquivalent(this VBScriptNameRewriter nameRewriter, NameToken x, NameToken y)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            return nameRewriter.GetMemberAccessTokenName(x) == nameRewriter.GetMemberAccessTokenName(y);
        }

        /// <summary>
        /// This is used by the GetMemberAccessTokenName for tokens that are not already NameToken instances. This derived type is used
        /// since it will bypass some of the the validation in the NameToken base constructor.
        /// </summary>
        private class ForRenamingNameToken : NameToken
        {
            public ForRenamingNameToken(string content, int lineIndex) : base(content, WhiteSpaceBehaviourOptions.Disallow, lineIndex) { }
        }
    }
}
