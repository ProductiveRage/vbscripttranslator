using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;
using VBScriptTranslator.LegacyParser.CodeBlocks.SourceRendering;
using VBScriptTranslator.LegacyParser.ContentBreaking;
using VBScriptTranslator.LegacyParser.Tokens;
using VBScriptTranslator.LegacyParser.Tokens.Basic;
using VBScriptTranslator.StageTwoParser.TokenCombining.NumberRebuilding;
using VBScriptTranslator.StageTwoParser.TokenCombining.OperatorCombinations;

namespace Tester
{
    static class Program
    {
        // =======================================================================================
        // ENTRY POINT
        // =======================================================================================
        static void Main()
        {
            (new VBScriptTranslator.UnitTests.StageTwoParser.OperatorCombinerTests()).OnePlusNegativeOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.OperatorCombinerTests()).OneMinusNegativeOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.OperatorCombinerTests()).OneMultipliedByPlusOne();

            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).NegativeOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).BracketedNegativeOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).PointOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).OnePointOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).NegativeOnePointOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).NegativePointOne();
            (new VBScriptTranslator.UnitTests.StageTwoParser.NumberRebuilderTests()).OnePlusNegativeOne();

            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.NonChangingTests()).DirectMethodWithNoArgumentsAndNoBrackets();
            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.NonChangingTests()).DirectMethodWithSingleArgumentWithBrackets();
            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.NonChangingTests()).ObjectMethodWithSingleArgumentWithBrackets();
            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.NonChangingTests()).DirectMethodWithSingleArgumentWithBracketsAndCallKeyword();
            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.NonChangingTests()).DirectMethodWithSingleArgumentWithBracketsAndCallKeywordAndExtraBracketsCall();

            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.ChangingTests()).DirectMethodWithSingleArgumentWithoutBrackets();
            (new VBScriptTranslator.UnitTests.LegacyParser.StatementBracketStandardisingTest.ChangingTests()).ObjectMethodWithSingleArgumentWithoutBrackets();

            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).StatementWithMemberAccessAndDecimalValueAndUnwrappedMethodArgument();
            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).StatementWithMemberAccessAndDecimalValueAndWrappedMethodArgument();
            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).SingleValueSetToNothing();
            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).TwoDimensionalArrayElementSetToNumber();
            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).TwoDimensionalArrayElementSetToNumberWithExplicitLet();
            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).TwoDimensionalArrayElementSetToNothing();
            (new VBScriptTranslator.UnitTests.LegacyParser.SingleStatementParsingTests()).TwoDimensionalArrayElementWithMethodCallIndexSetToNothing();

            var filename = "Test.vbs";

            Test_Old(filename);

            var statement1 = new Statement(
                new[]
                {
                    AtomToken.GetNewToken("WScript"),
                    AtomToken.GetNewToken("."),
                    AtomToken.GetNewToken("Echo"),
                    AtomToken.GetNewToken("("),
                    AtomToken.GetNewToken("1"),
                    AtomToken.GetNewToken("."),
                    AtomToken.GetNewToken("1"),
                    AtomToken.GetNewToken("+"),
                    new StringToken("2"),
                    AtomToken.GetNewToken(")")
                },
                Statement.CallPrefixOptions.Absent
            );
            var statement2 = new Statement(
                new[]
                {
                    AtomToken.GetNewToken("Go"),
                    AtomToken.GetNewToken("("),
                    AtomToken.GetNewToken(")")
                },
                Statement.CallPrefixOptions.Absent
            );
        }

        private static void Test_Old(string filename)
        {
            // Load raw script content
            var scriptContent = getScriptContent(filename);
            scriptContent = scriptContent.Replace("\r\n", "\n");

            // Break down content into String, Comment and UnprocessedContent tokens
            var tokens = StringBreaker.SegmentString(scriptContent);

            // Break down further into String, Comment, Atom and AbstractEndOfStatement tokens
            var atomTokens = new List<IToken>();
            foreach (var token in tokens)
            {
                if (token is UnprocessedContentToken)
                    atomTokens.AddRange(TokenBreaker.BreakUnprocessedToken((UnprocessedContentToken)token));
                else
                    atomTokens.Add(token);
            }

            // Translate these tokens into ICodeBlock implementations (representing
            // code VBScript structures)
            string[] endSequenceMet;
            var handler = new CodeBlockHandler(null);
            var codeBlocks = handler.Process(
                NumberRebuilder.Rebuild(OperatorCombiner.Combine(atomTokens)).ToList(),
                out endSequenceMet
            );

            // DEBUG: ender processed content as VBScript source code
            Console.WriteLine(getRenderedSourceVB(codeBlocks));
            Console.ReadLine();
        }

        // =======================================================================================
        // READ SCRIPT CONTENT
        // =======================================================================================
        private static string getScriptContent(string filename)
        {
            if ((filename ?? "").Trim() == "")
                throw new ArgumentException("getScriptContent: filename is null or blank");
            
            var file = new FileInfo(filename);
            using (var stream = file.OpenText())
            {
                return stream.ReadToEnd();
            }
        }

        // =======================================================================================
        // REBUILD ORIGINAL SOURCE FROM PROCESSED CONTENT
        // =======================================================================================
        private static string getRenderedSourceVB(List<ICodeBlock> content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            
            var output = new StringBuilder();
            foreach (var block in content)
                output.AppendLine(block.GenerateBaseSource(new SourceIndentHandler()));
            return output.ToString();
        }
    }
}
