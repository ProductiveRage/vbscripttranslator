using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks.Handlers;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.UnitTests.Shared.Comparers;
using Xunit;

namespace VBScriptTranslator.UnitTests.LegacyParser
{
    public class StatementHandlerTests
    {
        /// <summary>
        /// Only the first non-bracketed equality sign in a statement may indicate the separation between the value-to-set and the expression-to-set-it-to
        /// in a value-setting-statement, any subsequent equals signs are comparison operators (C# uses "==" for comparisons, as opposed to "=" for setting
        /// values, which is clearer.. but this is VBScript)
        /// </summary>
        [Fact]
        public void SubsequentEqualsTokensInValueSettingStatementAreComparisonOperators()
        {
            var statement = (new StatementHandler()).Process(new List<IToken>
            {
                new NameToken("bMatch", 0),
                new ComparisonOperatorToken("=", 0),
                new NumericValueToken("1", 0),
                new ComparisonOperatorToken("=", 0),
                new NumericValueToken("2", 0)
            });

            Assert.IsType<ValueSettingStatement>(statement);

            var valueSettingStatement = (ValueSettingStatement)statement;
            Assert.Equal(ValueSettingStatement.ValueSetTypeOptions.Let, valueSettingStatement.ValueSetType);
            Assert.Equal(
                new IToken[] { new NameToken("bMatch", 0) },
                valueSettingStatement.ValueToSet.Tokens,
                new TokenSetComparer()
            );
            Assert.Equal(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new ComparisonOperatorToken("=", 0),
                    new NumericValueToken("2", 0)
                },
                valueSettingStatement.Expression.Tokens,
                new TokenSetComparer()
            );
        }
    }
}
