using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class TokenSetComparer : IEqualityComparer<IEnumerable<IToken>>
    {
        public bool Equals(IEnumerable<IToken> x, IEnumerable<IToken> y)
        {
            if (x == null)
                throw new ArgumentNullException("tokensX");
            if (y == null)
                throw new ArgumentNullException("tokensY");

            var tokensArrayX = x.ToArray();
            var tokensArrayY = y.ToArray();
            if (tokensArrayX.Length != tokensArrayY.Length)
                return false;

            for (var index = 0; index < tokensArrayX.Length; index++)
            {
                var tokenX = tokensArrayX[index];
                var tokenY = tokensArrayY[index];
                var bytesX = GetBytes(tokenX);
                var bytesY = GetBytes(tokenY);
                if (bytesX.Length != bytesY.Length)
                    return false;
                for (var indexBytes = 0; indexBytes < bytesX.Length; indexBytes++)
                {
                    if (bytesX[indexBytes] != bytesY[indexBytes])
                        return false;
                }
            }
            return true;
        }

        private static byte[] GetBytes(object obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            using (var stream = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(stream, obj);
                stream.Seek(0, SeekOrigin.Begin);
                return ReadBytesFromStream(stream);
            }
        }

        private static byte[] ReadBytesFromStream(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            var buffer = new byte[4096];
            var read = 0;
            int chunk;
            while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
            {
                read += chunk;
                if (read == buffer.Length)
                {
                    int nextByte = stream.ReadByte();
                    if (nextByte == -1)
                        return buffer;
                    byte[] newBuffer = new byte[buffer.Length * 2];
                    Array.Copy(buffer, newBuffer, buffer.Length);
                    newBuffer[read] = (byte)nextByte;
                    buffer = newBuffer;
                    read++;
                }
            }
            var ret = new byte[read];
            Array.Copy(buffer, ret, read);
            return ret;
        }

        public int GetHashCode(IEnumerable<IToken> obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
