namespace VBScriptTranslator.CSharpWriter.CodeTranslation
{
	public enum ExpressionReturnTypeOptions
	{
		/// <summary>
		/// This refers to a true .net boolean, where the expression is going to be used in a C# if statement (VBScript always returns pseudo-booleans,
		/// represented by numbers, true and false are just -1 and 0, resp.)
		/// </summary>
		Boolean,

		/// <summary>
		/// This means that the expression is a non-value-returning statement, in some cases different logic is applied to statements than value-returning
		/// expressions and so a distinction is required between None and NotSpecified
		/// </summary>
		None,

		NotSpecified,

		/// <summary>
		/// This is used by SET statements, the return value must be an object reference (not a value type) otherwise it's an error condition
		/// </summary>
		Reference,

		/// <summary>
		/// This is used by non-SET allocation statements, the return value must not be an object reference (default members may be considered in order
		/// to return a value type from an object)
		/// </summary>
		Value
	}
}
