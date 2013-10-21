using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.UnitTests.Shared.Comparers;

namespace VBScriptTranslator.UnitTests.LegacyParser.Helpers
{
    public class CodeBlockComparer : IEqualityComparer<ICodeBlock>
    {
        public bool Equals(ICodeBlock x, ICodeBlock y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            if (x.GetType() != y.GetType())
                return false;

            var tokenSetComparer = new TokenSetComparer();

            if (x.GetType() == typeof(Statement))
                return tokenSetComparer.Equals(((Statement)x).Tokens, ((Statement)y).Tokens);
            else if (x.GetType() == typeof(ValueSettingStatement))
            {
                var valueSettingStatementX = (ValueSettingStatement)x;
                var valueSettingStatementY = (ValueSettingStatement)y;
                return (
                    (valueSettingStatementX.ValueSetType == valueSettingStatementY.ValueSetType) &&
					tokenSetComparer.Equals(valueSettingStatementX.ValueToSet.Tokens, valueSettingStatementY.ValueToSet.Tokens) &&
					tokenSetComparer.Equals(valueSettingStatementX.Expression.Tokens, valueSettingStatementY.Expression.Tokens)
                );
            }

            throw new NotSupportedException("Can not compare ICodeBlock of type " + x.GetType());
        }

        public int GetHashCode(ICodeBlock obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }

    }
}
