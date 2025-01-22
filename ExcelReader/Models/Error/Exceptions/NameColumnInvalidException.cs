using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class NameColumnInvalidException : Exception
    {
        public NameColumnInvalidException()
        {
        }

        public NameColumnInvalidException(string? message) : base(message)
        {
        }

        public NameColumnInvalidException(string? message, Exception? innerException) : base(message, innerException)
        {
        }

    }
}
