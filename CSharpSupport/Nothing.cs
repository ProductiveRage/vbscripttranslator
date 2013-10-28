namespace CSharpSupport
{
	/// <summary>
	/// This is like VBScript's non-value-type equivalent of Empty - internally we just a value that we can set to and compare to, but if this
	/// gets passed into a COM component then it should probably be replaced with null (true null, not VBScript's null) - TODO: Confirm
	/// </summary>
	public sealed class Nothing
	{
		// Singleton pattern courtesy of Jon Skeet ("Fifth version" from http://csharpindepth.com/Articles/General/Singleton.aspx)
		Nothing() { }

		public static Nothing Instance
		{
			get
			{
				return Nested.instance;
			}
		}

		private class Nested
		{
			// Explicit static constructor to tell C# compiler not to mark type as beforefieldinit
			static Nested() { }

			internal static readonly Nothing instance = new Nothing();
		}
	}
}
