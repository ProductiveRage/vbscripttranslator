using System;
using System.Linq;

namespace CSharpWriter.CodeTranslation
{
    public class CSharpName
    {
        public CSharpName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (name.Any(c => char.IsWhiteSpace(c)))
                throw new ArgumentException("Specified name contains whitespace - invalid");

            Name = name;
        }

        /// <summary>
        /// This will never be null, blank or contain any whitespace
        /// </summary>
        public string Name { get; private set; }
    }
}
