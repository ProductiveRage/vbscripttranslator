using System;

namespace VBScriptTranslator.RuntimeSupport.Attributes
{
	/// <summary>
	/// In order to fully implement VBScript TypeName support, we will need the original names of classes before they were changed for C# generation (for cases
	/// where they WERE changed). Generated classes should be decorated with this attribute to expose that information.
	/// </summary>
    public class SourceClassName : Attribute
    {
		public SourceClassName(string name)
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
