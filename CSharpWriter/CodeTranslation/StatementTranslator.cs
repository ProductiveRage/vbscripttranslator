using CSharpWriter.Misc;
using System;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.ExpressionParsing;
using LegacyParser = VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation
{
    public class StatementTranslator : ITranslateIndividualStatements
    {
        private readonly CSharpName _supportClassName;
        private readonly VBScriptNameRewriter _nameRewriter;
        private readonly TempValueNameGenerator _tempNameGenerator;
        public StatementTranslator(CSharpName supportClassName, VBScriptNameRewriter nameRewriter, TempValueNameGenerator tempNameGenerator)
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
        /// This will never return null or blank, it will raise an exception if unable to satisfy the request (this includes the case of
        /// a null statement reference)
        /// </summary>
        public string Translate(LegacyParser.Statement statement)
        {
            if (statement == null)
                throw new ArgumentNullException("statement");

            var expressions = ExpressionGenerator.Generate(statement.BracketStandardisedTokens).ToArray();
            if (expressions.Length != 1)
                throw new ArgumentException("Statement translation should always result in a single expression being generated");

            return Translate(expressions[0]);
        }

        private string Translate(Expression expression)
        {
            if (expression == null)
                throw new ArgumentNullException("expression");

            // Assert expectations about numbers of segments and operators (if any)
            var segments = expression.Segments.ToArray();
            if (segments.Length == 0)
                throw new ArgumentException("The expression was broken down into zero segments - invalid content");
            if (segments.Length > 3)
                throw new ArgumentException("Expressions with more than three segments are invalid (they must be broken down further), this one has " + segments.Length);
            var operatorSegmentsWithIndexes = segments
                .Select((segment, index) => new { Segment = segment as OperationExpressionSegment, Index = index })
                .Where(s => s.Segment != null);
            if (operatorSegmentsWithIndexes.Count() > 1)
                throw new ArgumentException("Expressions with more than one operators are invalid (they must be broken down further), this one has " + operatorSegmentsWithIndexes.Count());

            // Assert expectations about combinations of segments and operators (if any)
            var operatorSegmentWithIndex = operatorSegmentsWithIndexes.SingleOrDefault();
            if (operatorSegmentWithIndex == null)
            {
                if (segments.Length != 1)
                    throw new ArgumentException("Expressions with multiple segments are not invalid if there is no operator");
            }
            else
            {
                if (segments.Length == 1)
                    throw new ArgumentException("Expressions containing only a single segment, where that segment is an operator, are invalid");
                else if (segments.Length == 2)
                {
                    if (operatorSegmentWithIndex.Index != 0)
                        throw new ArgumentException("If there are any only two segments then the first must be a negation operator (the first here isn't an operator)");
                    if ((operatorSegmentWithIndex.Segment.Token.Content != "-")
                    && !operatorSegmentWithIndex.Segment.Token.Content.Equals("NOT", StringComparison.InvariantCultureIgnoreCase))
                        throw new ArgumentException("If there are any only two segments then the first must be a negation operator (here it has the token content \"" + operatorSegmentWithIndex.Segment.Token.Content + "\")");
                }
                else if (operatorSegmentWithIndex.Index != 1)
                    throw new ArgumentException("If there are three segments, then the middle must be an operator");
            }

            if (segments.Length == 1)
                return TranslateNonOperatorSegment(segments[0]);

            if (segments.Length == 2)
            {
                return string.Format(
                    "{0}.{1}({2})",
                    _supportClassName.Name,
                    GetSupportFunctionName(operatorSegmentWithIndex.Segment.Token),
                    TranslateNonOperatorSegment(segments[1])
                );
            }

            return string.Format(
                "{0}.{1}({2}, {3})",
                _supportClassName.Name,
                TranslateNonOperatorSegment(segments[0]),
                GetSupportFunctionName(operatorSegmentWithIndex.Segment.Token),
                TranslateNonOperatorSegment(segments[2])
            );
        }

        private string TranslateNonOperatorSegment(IExpressionSegment segment)
        {
            if (segment == null)
                throw new ArgumentNullException("segment");
            if (segment is OperationExpressionSegment)
                throw new ArgumentException("This will not accept OperationExpressionSegment instances");

            var numericValueSegment = segment as NumericValueExpressionSegment;
            if (numericValueSegment != null)
                return numericValueSegment.Token.Content;

            var stringValueSegment = segment as StringValueExpressionSegment;
            if (stringValueSegment != null)
                return stringValueSegment.Token.Content.ToLiteral();

            var callExpressionSegment = segment as CallExpressionSegment;
            if (callExpressionSegment != null)
                return Translate(callExpressionSegment);

            var bracketedExpressionSegment = segment as BracketedExpressionSegment;
            if (bracketedExpressionSegment != null)
                return Translate(bracketedExpressionSegment);

            throw new NotImplementedException(); // TODO
        }

        private string Translate(BracketedExpressionSegment bracketedExpressionSegment)
        {
            if (bracketedExpressionSegment == null)
                throw new ArgumentNullException("bracketedExpressionSegment");

            throw new NotImplementedException(); // TODO
        }

        private string Translate(CallExpressionSegment callExpressionSegment)
        {
            if (callExpressionSegment == null)
                throw new ArgumentNullException("callExpressionSegment");

            // The "master" CALL method signature is
            //
            //   CALL(object target, IEnumerable<string> members, IEnumerable<object> arguments)
            //
            // but there are alternate signatures to try to make the most common calls easier to read - eg.
            //
            //   CALL(object)
            //   CALL(object, params object[] arguments)
            //   CALL(object, string member1, params object[] arguments)
            //   CALL(object, string member1, string member2, params object[] arguments)
            //   ..
            //   CALL(object, string member1, string member2, string member3, string member4, string member5, params object[] arguments)
            //
            // The maximum number of member access tokens (one being the initial object and further being the properties / functions accessed)
            // that may use one of these alternate signatures is stored in the constant maxNumberOfMemberAccessorBeforeArraysRequired
            const int maxNumberOfMemberAccessorBeforeArraysRequired = 6;

            var callExpressionContent = new StringBuilder();

            // Note: Even if there is only a single member access token (meaning call for "a" not "a.Test") and there are no arguments, we
            // still need to use the CALL method to account for any handling of default properties
            var firstMemberAccessToken = callExpressionSegment.MemberAccessTokens.First();
            callExpressionContent.AppendFormat(
                "{0}.CALL({1}",
                _supportClassName.Name,
                GetMemberAccessTokenName(firstMemberAccessToken)
            );

            var numberOfAccessTokens = callExpressionSegment.MemberAccessTokens.Count();
            if (numberOfAccessTokens > 1)
            {
                callExpressionContent.Append(", ");
                if (numberOfAccessTokens > maxNumberOfMemberAccessorBeforeArraysRequired)
                    callExpressionContent.Append(" new[] { ");
                for (var index = 1; index < numberOfAccessTokens; index++)
                {
                    callExpressionContent.Append(
                        GetMemberAccessTokenName(callExpressionSegment.MemberAccessTokens.ElementAt(index))
                    );
                    if (index < (numberOfAccessTokens - 1))
                        callExpressionContent.Append(", ");
                }
                if (numberOfAccessTokens > maxNumberOfMemberAccessorBeforeArraysRequired)
                    callExpressionContent.Append(" }");
            }

            if (callExpressionSegment.Arguments.Any())
            {
                callExpressionContent.Append(", ");
                if (numberOfAccessTokens > maxNumberOfMemberAccessorBeforeArraysRequired)
                    callExpressionContent.Append("new object[] { ");
                var numberOfArguments = callExpressionSegment.Arguments.Count();
                for (var index = 0; index < numberOfArguments; index++)
                {
                    callExpressionContent.Append(
                        Translate(callExpressionSegment.Arguments.ElementAt(index))
                    );
                    if (index < (numberOfArguments - 1))
                        callExpressionContent.Append(", ");
                }
                if (numberOfAccessTokens > maxNumberOfMemberAccessorBeforeArraysRequired)
                    callExpressionContent.Append(" }");
            }

            callExpressionContent.Append(")");
            return callExpressionContent.ToString();
        }

        private string GetSupportFunctionName(OperatorToken operatorToken)
        {
            if (operatorToken == null)
                throw new ArgumentNullException("operatorToken");

            switch (operatorToken.Content.ToUpper())
            {
                // Arithmetic operators
                case "^":   return "POW";
                case "/":   return "DIV";
                case "*":   return "MULT";
                case "\\":  return "INTDIV";
                case "MOD": return "MOD";
                case "+":   return "ADD";
                case "-":   return "SUBT";

                // String concatenation
                case "&":   return "CONCAT";

                // Logical operators
                case "NOT": return "NOT";
                case "AND": return "AND";
                case "OR":  return "OR";
                case "XOR": return "XOR";

                // Comparison operators
                case "=":   return "EQ";
                case "<>":  return "NOTEQ";
                case "<":   return "LT";
                case ">":   return "GT";
                case "<=":  return "LTE";
                case ">=":  return "GTE";
                case "IS":  return "IS";
                case "EQV": return "EQV";
                case "IMP": return "IMP";

                // Error
                default:
                    throw new NotSupportedException("Unsupported OperatorToken content: " + operatorToken.Content);
            }
        }

        /// <summary>
        /// When trying to access variables, functions, classes, etc.. we need to pass 
        /// TODO: Explain..
        /// </summary>
        private string GetMemberAccessTokenName(IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            var nameToken = (token as NameToken) ?? new ForRenamingNameToken(token.Content);
            return _nameRewriter(nameToken).Name.ToLiteral();
        }

        /// <summary>
        /// TODO: Explain..
        /// </summary>
        private class ForRenamingNameToken : NameToken
        {
            public ForRenamingNameToken(string content) : base(content, WhiteSpaceBehaviourOptions.Disallow) { }
        }
    }
}
