using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
     public class InvalidRowException : Exception
    {
        public InvalidRowException()
        {
        }

        public InvalidRowException(string? message) : base(message)
        {
        }

        public InvalidRowException(string? message, Exception? innerException) : base(message, innerException)
        {
        }

    }
}
