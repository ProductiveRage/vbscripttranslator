using System;
using System.Collections.Generic;
using System.Text;
using CSharpWriter;
using CSharpWriter.CodeTranslation;
using CSharpWriter.CodeTranslation.BlockTranslators;
using CSharpWriter.Lists;

namespace Tester
{
    public static class DefaultASPTranslator
    {
        public static NonNullImmutableList<TranslatedStatement> Translate(string aspScriptContent, bool renderCommentsAboutUndeclaredVariables = true)
        {
            if (aspScriptContent == null)
                throw new ArgumentNullException("aspScriptContent");

            return DefaultTranslator.Translate(
                TransformIntoStandardScript(aspScriptContent),
                new[] { "Application", "Response", "Request", "Session", "Server" }, // Assume identify these expected references as undeclared variables
                OuterScopeBlockTranslator.OutputTypeOptions.Executable,
                renderCommentsAboutUndeclaredVariables
            );
        }

        private static string TransformIntoStandardScript(string aspScriptContent)
        {
            if (aspScriptContent == null)
                throw new ArgumentNullException("aspScriptContent");

            var standardScriptContent = new StringBuilder();
            while (true)
            {
                var indexOfStartOfNextScriptBlock = aspScriptContent.IndexOf("<%");
                if (indexOfStartOfNextScriptBlock == -1)
                    break;

                if (indexOfStartOfNextScriptBlock > 0)
                    standardScriptContent.AppendLine(GetResponseWriteForStaticContent(aspScriptContent.Substring(0, indexOfStartOfNextScriptBlock)));

                var indexOfEndOfScriptBlock = aspScriptContent.IndexOf("%>", indexOfStartOfNextScriptBlock);
                var scriptBlockContent = aspScriptContent.Substring(indexOfStartOfNextScriptBlock + 2, (indexOfEndOfScriptBlock - indexOfStartOfNextScriptBlock) - 2).Trim();
                if (scriptBlockContent.StartsWith("="))
                    scriptBlockContent = "Response.Write " + scriptBlockContent.Substring(1);
                standardScriptContent.AppendLine(scriptBlockContent);
                aspScriptContent = aspScriptContent.Substring(indexOfEndOfScriptBlock + 2);
            }
            if (aspScriptContent.Length > 0)
                standardScriptContent.AppendLine(GetResponseWriteForStaticContent(aspScriptContent));
            return standardScriptContent.ToString();
        }

        private static string GetResponseWriteForStaticContent(string content)
        {
            if (string.IsNullOrEmpty(content))
                throw new ArgumentException("content may not be null or blank (it may be whitespace-only, since that may have significance for the rendered markup)");

            var stringsToRender = new List<string>();
            var stringContentBuffer = new StringBuilder();
            for (var index = 0; index < content.Length; index++)
            {
                // See http://www.devguru.com/technologies/vbscript/13893
                var character = content[index];
                string controlCharacterContentIfAny;
                if (character == 13)
                {
                    if ((index < (content.Length - 1)) && (content[index + 1] == 10))
                    {
                        if ((index < (content.Length - 2)) && (content[index + 2] == 10))
                        {
                            controlCharacterContentIfAny = "vbNewLine";
                            index += 2;
                        }
                        else
                        {
                            controlCharacterContentIfAny = "vbCrLf";
                            index += 1;
                        }
                    }
                    else
                        controlCharacterContentIfAny = "vbCr";
                }
                else if (character == 10)
                    controlCharacterContentIfAny = "vbLF";
                else if (character == 9)
                    controlCharacterContentIfAny = "vbTab";
                else if (character == 11)
                    controlCharacterContentIfAny = "vbVerticalTab";
                else if (character == 11)
                    controlCharacterContentIfAny = "vbFormFeed";
                else if (character == 0)
                    controlCharacterContentIfAny = "vbNullChar";
                else
                {
                    // Not a control character
                    stringContentBuffer.Append(character);
                    continue;
                }

                if (stringContentBuffer.Length > 0)
                {
                    stringsToRender.Add("\"" + stringContentBuffer.ToString().Replace("\"", "\"\"") + "\""); // TODO: Explain double-quoting
                    stringContentBuffer.Clear();
                }
                stringsToRender.Add(controlCharacterContentIfAny);
            }
            if (stringContentBuffer.Length > 0)
                stringsToRender.Add("\"" + stringContentBuffer.ToString().Replace("\"", "\"\"") + "\""); // TODO: Explain double-quoting (and remove duplication?)

            return "Response.Write " + string.Join(" & ", stringsToRender);
        }
    }
}
