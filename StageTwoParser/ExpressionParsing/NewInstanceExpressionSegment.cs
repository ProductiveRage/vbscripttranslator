using System;
using System.Collections.Generic;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace VBScriptTranslator.StageTwoParser.ExpressionParsing
{
    public class NewInstanceExpressionSegment : IExpressionSegment
    {
        public NewInstanceExpressionSegment(NameToken className)
        {
            if (className == null)
                throw new ArgumentNullException("className");

            ClassName = className;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken ClassName { get; private set; }

		/// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		IEnumerable<IToken> IExpressionSegment.AllTokens
		{
			get
			{
                return new IToken[]
                {
                    new KeyWordToken("new", ClassName.LineIndex),
                    ClassName
                };
			}
		}

        public string RenderedContent
        {
            get { return "new " + ClassName.Content; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
