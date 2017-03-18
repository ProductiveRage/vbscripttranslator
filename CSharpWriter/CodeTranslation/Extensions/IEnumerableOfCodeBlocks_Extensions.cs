using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.CSharpWriter.Lists;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions
{
	public static class IEnumerableOfCodeBlocks_Extensions
	{
		public static bool DoesScopeContainOnErrorResumeNext(this IEnumerable<ICodeBlock> blocks)
		{
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			foreach (var block in blocks)
			{
				if (block == null)
					throw new ArgumentException("Null reference encountered in blocks set");

				if (block is OnErrorResumeNext)
					return true;

				if (block is IDefineScope)
				{
					// If this block defines its own scope then it doesn't matter if it contains an OnErrorResumeNext statement
					// within it, it won't affect this scope
					continue;
				}

				var blockWithNestedContent = block as IHaveNestedContent;
				if ((blockWithNestedContent != null) && blockWithNestedContent.AllExecutableBlocks.ToNonNullImmutableList().DoesScopeContainOnErrorResumeNext())
					return true;
			}
			return false;
		}

		public static IEnumerable<IToken> EnumerateAllTokens(this IEnumerable<ICodeBlock> blocks)
		{
			if (blocks == null)
				throw new ArgumentNullException("blocks");

			foreach (var block in blocks)
			{
				if (block == null)
					throw new ArgumentException("Null reference encountered in blocks set");

				IEnumerable<Statement> expressionsToInterrogate;
				var nonNestedExpressionContainingBlock = block as IHaveNonNestedExpressions;
				if (nonNestedExpressionContainingBlock != null)
					expressionsToInterrogate = nonNestedExpressionContainingBlock.NonNestedExpressions;
				else
				{
					var statement = block as Statement;
					if (statement != null)
						expressionsToInterrogate = new[] { statement };
					else
						expressionsToInterrogate = new Statement[0];
				}
				foreach (var token in expressionsToInterrogate.SelectMany(e => e.Tokens))
					yield return token;

				var nestedContentBlock = block as IHaveNestedContent;
				if (nestedContentBlock != null)
				{
					foreach (var nestedToken in EnumerateAllTokens(nestedContentBlock.AllExecutableBlocks))
						yield return nestedToken;
				}
			}
		}
	}
}
