using System;

namespace Tester
{
    public class ResponseMock
    {
        public void Write(object content)
        {
            Console.Write(content);
        }
    }
}
