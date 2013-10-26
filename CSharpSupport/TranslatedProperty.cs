using System;

namespace CSharpSupport
{
	/// <summary>
    /// Since C# doesn't support named index properties, where these exist in VBScript source they are converted into methods for the get and set (depending
    /// upon which are present in the source) with this attribute. The getter and setter will have non-void and void return types, resp.
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
