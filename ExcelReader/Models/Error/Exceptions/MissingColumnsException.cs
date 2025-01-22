using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class MissingColumnsException : Exception
    {
        public MissingColumnsException()
        {
        }

        public MissingColumnsException(string? message) : base(message)
        {
        }

        public MissingColumnsException(string? message, Exception? innerException) : base(message, innerException)
        {
        }

    }
}
