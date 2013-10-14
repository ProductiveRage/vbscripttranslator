using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using VBScriptTranslator.LegacyParser.Tokens;

namespace VBScriptTranslator.UnitTests.Shared.Comparers
{
    public class TokenComparer : IEqualityComparer<IToken>
    {
        public bool Equals(IToken x, IToken y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var bytesX = GetBytes(x);
            var bytesY = GetBytes(y);
            if (bytesX.Length != bytesY.Length)
                return false;
            for (var indexBytes = 0; indexBytes < bytesX.Length; indexBytes++)
            {
                if (bytesX[indexBytes] != bytesY[indexBytes])
                    return false;
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

        public int GetHashCode(IToken obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
