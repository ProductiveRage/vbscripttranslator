using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions
{
	/// <summary>
	/// This is a special derived class of NameToken, it will not be affected when passed through the GetMemberAccessTokenName extension method of a VBScriptNameRewriter
	/// (this may be useful when content is being injected into expressions to ensure that name rewriting isn't double-applied - it is used in the StatementTranslator,
	/// for example)
	/// </summary>
    [Serializable]
    public class DoNotRenameNameToken : NameToken
	{
		public DoNotRenameNameToken(string content, int lineIndex) : base(content, WhiteSpaceBehaviourOptions.Allow, lineIndex)
		{
			if (string.IsNullOrWhiteSpace(content))
				throw new ArgumentException("Null/blank content specified");
		}
	}
}
