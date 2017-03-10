using System;

namespace VBScriptTranslator.RuntimeSupport.Attributes
{
	/// <summary>
	/// Since C# doesn't support named index properties, where these exist in the VBScript source they are converted into methods for the get and/or set (depending
	/// upon which are present in the source) with this attribute. To avoid confusion (why does this from-property method have the attribute and this one not?),
	/// this attribute may be applied to all methods that are translated from VBScript properties. The getter and setter of this methods will have non-void
	/// and void return types, resp.
	/// </summary>
	public class TranslatedProperty : Attribute
	{
		public TranslatedProperty(string name)
		{
			if (name == null)
				throw new ArgumentNullException("name");

			Name = name;
		}

		/// <summary>
		/// This will never be null (but this is pretty much the only guarantee we can make due to VBScript's crazy variable name escaping support)
		/// </summary>
		public string Name { get; private set; }
	}
}
