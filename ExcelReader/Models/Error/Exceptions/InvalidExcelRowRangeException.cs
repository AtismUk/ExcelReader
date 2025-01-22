using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class InvalidExcelRowRangeException : Exception
    {
        public InvalidExcelRowRangeException()
        {
        }

        public InvalidExcelRowRangeException(string message)
            : base(message)
        {
        }

        public InvalidExcelRowRangeException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
