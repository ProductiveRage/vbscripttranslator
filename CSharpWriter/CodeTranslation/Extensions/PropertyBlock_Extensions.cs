using System;
using System.Linq;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace VBScriptTranslator.CSharpWriter.CodeTranslation.Extensions
{
    public static class PropertyBlock_Extensions
    {
        public static bool IsIndexedProperty(this PropertyBlock source)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if ((source.PropType == PropertyBlock.PropertyType.Get) && source.Parameters.Any())
                return true;

            return source.Parameters.Count() > 1;
        }
    }
}
