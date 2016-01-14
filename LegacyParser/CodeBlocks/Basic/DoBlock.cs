using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;

namespace VBScriptTranslator.LegacyParser.CodeBlocks.Basic
{
	[Serializable]
	public class DoBlock : ILoopOverNestedContent, ICodeBlock
	{
		// =======================================================================================
		// CLASS INITIALISATION
		// =======================================================================================
		/// <summary>
		/// It is valid to have a null conditionStatement in VBScript - in case the isPreCondition and doUntil constructor arguments are of no consequence
		/// </summary>
		public DoBlock(Expression conditionIfAny, bool isPreCondition, bool doUntil, bool supportsExit, IEnumerable<ICodeBlock> statements, int lineIndexOfStartOfConstruct)
		{
			if (statements == null)
				throw new ArgumentNullException("statements");
			if (lineIndexOfStartOfConstruct < 0 )
				throw new ArgumentOutOfRangeException("Must be zero or greater", "lineIndexOfStartOfConstruct");

			ConditionIfAny = conditionIfAny;
			IsPreCondition = isPreCondition;
			IsDoWhileCondition = !doUntil;
			SupportsExit = supportsExit;
			Statements = statements.ToList().AsReadOnly();
			if (Statements.Any(s => s == null))
				throw new ArgumentException("Null reference encountered in statements set");
			LineIndexOfStartOfConstruct = lineIndexOfStartOfConstruct;
		}

		// =======================================================================================
		// PUBLIC DATA ACCESS
		// =======================================================================================
		/// <summary>
		/// This may be null since VBScript supports DO..WHILE loops with no constraint
		/// </summary>
		public Expression ConditionIfAny { get; private set; }

		public bool IsPreCondition { get; private set; }

		/// <summary>
		/// If this is true then for construct is of the for DO WHILE..LOOP or DO..LOOP WHILE, as opposed to DO UNTIL..LOOP or DO..LOOP UNTIL
		/// </summary>
		public bool IsDoWhileCondition { get; private set; }

		/// <summary>
		/// If this represents a WHILE construct then the source VBScript does not allow it to be targeted by an EXIT DO statement (and there is
		/// no EXIT WHILE), so this will be false. If this represent a DO construct then this will be true.
		/// </summary>
		public bool SupportsExit { get; private set; }

		/// <summary>
		/// This will never be null nor contain any null references, but it may be an empty set
		/// </summary>
		public IEnumerable<ICodeBlock> Statements { get; private set; }

		/// <summary>
		/// This will always be zero or greater
		/// </summary>
		public int LineIndexOfStartOfConstruct { get; private set; }

		/// <summary>
		/// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
		/// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
		/// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
		/// blocks within child code blocks.
		/// </summary>
		IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
		{
			get
			{
				var statementBlocks = Statements.ToList();
				if (ConditionIfAny != null)
				{
					if (IsPreCondition)
						statementBlocks.Insert(0, ConditionIfAny);
					else
						statementBlocks.Add(ConditionIfAny);
				}
				return statementBlocks;
			}
		}

		// =======================================================================================
		// VBScript BASE SOURCE RE-GENERATION
		// =======================================================================================
		/// <summary>
		/// Re-generate equivalent VBScript source code for this block - there should not be a line return at the end of the content
		/// </summary>
		public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
		{
			var output = new StringBuilder();

			// Open statement (with condition if this construct has a pre condition)
			output.Append(indenter.Indent + "Do");
			if (IsPreCondition && (ConditionIfAny != null))
			{
				output.Append(" ");
				if (IsDoWhileCondition)
					output.Append("While ");
				else
					output.Append("Until ");
				output.AppendLine(ConditionIfAny.GenerateBaseSource(NullIndenter.Instance));
			}
			else
				output.AppendLine();

			// Render inner content
			foreach (var statement in Statements)
				output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

			// Close statement (with condition if this construct has a pre condition)
			output.Append(indenter.Indent + "Loop");
			if (!IsPreCondition && (ConditionIfAny != null))
			{
				output.Append(" ");
				if (IsDoWhileCondition)
					output.Append("While ");
				else
					output.Append("Until ");
				output.Append(ConditionIfAny.GenerateBaseSource(NullIndenter.Instance));
			}
			return output.ToString();
		}
	}
}
