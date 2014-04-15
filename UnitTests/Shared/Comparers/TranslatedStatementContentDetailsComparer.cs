using CSharpWriter.CodeTranslation;
using System;
using System.Collections.Generic;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
	public class TranslatedStatementContentDetailsComparer : IEqualityComparer<TranslatedStatementContentDetails>
	{
		public bool Equals(TranslatedStatementContentDetails x, TranslatedStatementContentDetails y)
		{
			if (x == null)
				throw new ArgumentNullException("x");
			if (y == null)
				throw new ArgumentNullException("y");

			if (x.TranslatedContent != y.TranslatedContent)
				return false;

			var tokenSetComparer = new TokenSetComparer();
			return tokenSetComparer.Equals(x.VariablesAccessed, y.VariablesAccessed);
		}

		public int GetHashCode(TranslatedStatementContentDetails obj)
		{
			if (obj == null)
				throw new ArgumentNullException("obj");

			return 0;
		}
	}
}
