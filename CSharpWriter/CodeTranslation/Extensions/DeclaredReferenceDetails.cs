using System;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public class DeclaredReferenceDetails
    {
        public DeclaredReferenceDetails(ReferenceTypeOptions referenceType, ScopeLocationOptions scopeLocation)
        {
            if (!Enum.IsDefined(typeof(ReferenceTypeOptions), referenceType))
                throw new ArgumentOutOfRangeException("referenceType");
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");

            ReferenceType = referenceType;
            ScopeLocation = scopeLocation;
        }

        public ReferenceTypeOptions ReferenceType { get; private set; }
        public ScopeLocationOptions ScopeLocation { get; private set; }
    }
}
