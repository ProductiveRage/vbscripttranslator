using System;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace CSharpSupport
{
    public class TranslatedPropertyIReflectImplementation : BasicIReflectImplementation
	{
        // TODO: This needs to intercept requests for indexed properties and pass them through to the method with the appropriate TranslatedProperty attribute
	}
}
