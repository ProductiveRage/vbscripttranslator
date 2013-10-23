using CSharpSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public class VBScriptEsqueValueRetrieverTests
    {
        [Fact]
        public void VALDoesNotAlterMinusOneNull()
        {
            Assert.Equal(
                null,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(null)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOne()
        {
            Assert.Equal(
                1,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneFloat()
        {
            Assert.Equal(
                1.1f,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1.1f)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneDouble()
        {
            Assert.Equal(
                1.1d,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1.1d)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneDecimal()
        {
            Assert.Equal(
                1.1m,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1.1m)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOne()
        {
            Assert.Equal(
                -1,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(-1)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusEmptyString()
        {
            Assert.Equal(
                "",
                (new VBScriptEsqueValueRetriever(name => name)).VAL("")
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusNonEmptyString()
        {
            Assert.Equal(
                "Test",
                (new VBScriptEsqueValueRetriever(name => name)).VAL("Test")
            );
        }

        [Fact]
        public void IFOfNullIsFalse()
        {
            Assert.False(
                (new VBScriptEsqueValueRetriever(name => name)).IF(null)
            );
        }

        [Fact]
        public void IFOfZeroIsFalse()
        {
            Assert.False(
                (new VBScriptEsqueValueRetriever(name => name)).IF(0)
            );
        }

        [Fact]
        public void IFOfOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(1)
            );
        }

        [Fact]
        public void IFOfMinusOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(-1)
            );
        }

        [Fact]
        public void IFOfOnePointOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(1.1)
            );
        }

        /// <summary>
        /// VBScript doesn't round the number down to zero and find 0.1 to be false, it just checks that the number is non-zero
        /// </summary>
        [Fact]
        public void IFOfPointOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(0.1)
            );
        }

        [Fact]
        public void IFOfPointStringRepresentationOfOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF("1")
            );
        }

        [Fact]
        public void IFThrowsExceptionForBlanksString()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                (new VBScriptEsqueValueRetriever(name => name)).IF("");
            });
        }

        [Fact]
        public void IFThrowsExceptionForNonNumericString()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                (new VBScriptEsqueValueRetriever(name => name)).IF("one");
            });
        }

        [Fact]
        public void IFIgnoresWhiteSpaceWhenParseingStrings()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF("   1    ")
            );
        }
    }
}
