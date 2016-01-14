using VBScriptTranslator.RuntimeSupport.Attributes;

namespace VBScriptTranslator.UnitTests.RuntimeSupport.Implementations
{
    /// <summary>
    /// This is an example of the type of class that may be emitted by the translation process, one with a parameter-less default member
    /// </summary>
    [SourceClassName("ExampleDefaultPropertyType")]
    public class exampledefaultpropertytype
    {
        [IsDefault]
        public object result { get; set; }
    }
}
