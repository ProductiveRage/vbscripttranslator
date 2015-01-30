using System;
using System.Collections.Generic;
using System.Linq;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class RuntimeErrorExpressionSegment : IExpressionSegment
    {
        public RuntimeErrorExpressionSegment(string originalContent, IEnumerable<IToken> originalTokens, Type exceptionType, string message)
        {
            if (string.IsNullOrWhiteSpace(originalContent))
                throw new ArgumentException("Null/blank originalContent specified");
            if (originalTokens == null)
                throw new ArgumentNullException("originalTokens");
            if (exceptionType == null)
                throw new ArgumentNullException("exceptionType");
            if (string.IsNullOrWhiteSpace(message))
                throw new ArgumentException("Null/blank message specified");

            AllTokens = originalTokens.ToList().AsReadOnly();
            if (!AllTokens.Any())
                throw new ArgumentException("Empty originalTokens set specified - invalid");
            if (AllTokens.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in originalTokens set - invalid");

            if (!typeof(Exception).IsAssignableFrom(exceptionType))
                throw new ArgumentException("The specified exceptionType must be derived from Exception");
            var acceptableConstructors = exceptionType.GetConstructors()
                .Where(c =>
                {
                    var constructorParameters = c.GetParameters();
                    if (!constructorParameters.Any() || (constructorParameters.First().ParameterType != typeof(string)))
                        return false;
                    return (constructorParameters.Count() == 1) || constructorParameters.Skip(1).All(p => p.IsOptional);
                });
            if (!acceptableConstructors.Any())
            if (exceptionType.GetConstructor(new[] { typeof(string) }) == null)
                throw new ArgumentException("The specified exceptionType must have a public constructor that takes a single string argument");

            RenderedContent = originalContent;
            ExceptionType = exceptionType;
            Message = message;
        }

        /// <summary>
        /// This will never be null or blank
        /// </summary>
        public string RenderedContent { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        public IEnumerable<IToken> AllTokens { get; private set; }

        /// <summary>
        /// This will always be a type that is derived from Exception and that has a public constructor that takes a single string argument
        /// </summary>
        public Type ExceptionType { get; private set; }

        /// <summary>
        /// This will never be null or blank
        /// </summary>
        public string Message { get; private set; }
    }
}
