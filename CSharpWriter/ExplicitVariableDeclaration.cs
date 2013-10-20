using System;
using VBScriptTranslator.LegacyParser.Tokens.Basic;

namespace CSharpWriter
{
    public class VariableDeclaration
    {
        public VariableDeclaration(NameToken name, VariableDeclarationScopeOptions scope, bool isArray)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!Enum.IsDefined(typeof(VariableDeclarationScopeOptions), scope))
                throw new ArgumentOutOfRangeException("scope");

            Name = name;
            Scope = scope;
            IsArray = isArray;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken Name { get; private set; }

        public VariableDeclarationScopeOptions Scope { get; private set; }

        public bool IsArray { get; private set; }
    }
}
