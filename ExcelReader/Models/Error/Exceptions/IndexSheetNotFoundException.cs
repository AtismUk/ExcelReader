using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class IndexSheetNotFoundException : Exception
    {
        public IndexSheetNotFoundException()
        {
        }

        public IndexSheetNotFoundException(string message)
            : base(message)
        {
        }

        public IndexSheetNotFoundException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
