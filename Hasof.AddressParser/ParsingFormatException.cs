using System;

namespace Hasof.AddressParser
{
    public class ParsingFormatException : Exception
    {
        public ParsingFormatException(string message) : base(message) { }
    }
}