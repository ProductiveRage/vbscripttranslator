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
        public static string GetMemberAccessTokenName(this VBScriptNameRewriter source, IToken token)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (token == null)
                throw new ArgumentNullException("token");

            var nameToken = (token as NameToken) ?? new ForRenamingNameToken(token.Content, token.LineIndex);
			if (nameToken is DoNotRenameNameToken)
				return nameToken.Content;
            return source(nameToken).Name;
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
