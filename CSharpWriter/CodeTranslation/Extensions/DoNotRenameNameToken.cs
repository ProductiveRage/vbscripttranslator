using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
	/// <summary>
	/// This is a special derived class of NameToken, it will not be affected when passed through the GetMemberAccessTokenName extension method of a VBScriptNameRewriter
	/// (this may be useful when content is being injected into expressions to ensure that name rewriting isn't double-applied - it is used in the StatementTranslator,
	/// for example)
	/// </summary>
	public class DoNotRenameNameToken : NameToken
	{
		public DoNotRenameNameToken(string content) : base(content, WhiteSpaceBehaviourOptions.Allow)
		{
			if (string.IsNullOrWhiteSpace(content))
				throw new ArgumentException("Null/blank content specified");
		}
	}
}
