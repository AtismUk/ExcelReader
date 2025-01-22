using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class DuplicateColumnException : Exception
    {
        public DuplicateColumnException()
        {
        }

        public DuplicateColumnException(string message)
            : base(message)
        {
        }

        public DuplicateColumnException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
