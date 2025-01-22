using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelReader.Models
{
    public class TestModel
    {
        public string Id { get; set; }

        public PersonModel Person { get; set; } = new();

        public DocumentModel Document { get; set; } = new();

        public int? TestNull { get; set; }
    }
}
