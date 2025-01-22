using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Error.Exceptions
{
    internal class PropertyRequiredException : Exception
    {
        public PropertyRequiredException()
        {
        }

        public PropertyRequiredException(string message)
            : base(message)
        {
        }

        public PropertyRequiredException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
