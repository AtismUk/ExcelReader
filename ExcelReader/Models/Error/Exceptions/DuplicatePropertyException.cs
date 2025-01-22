using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class DuplicatePropertyException : Exception
    {
        public DuplicatePropertyException()
        {
        }

        public DuplicatePropertyException(string message)
            : base(message)
        {
        }

        public DuplicatePropertyException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
