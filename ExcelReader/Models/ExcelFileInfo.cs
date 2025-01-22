using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models
{
    public class ExcelFileInfo
    {
        public string Name { get; }
        public string Path { get; }

        public ExcelFileInfo(string name, string path)
        {
            Name = name;
            Path = path;
        }
    }
}
