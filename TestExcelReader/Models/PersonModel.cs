using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelReader.Models
{
    public class PersonModel
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public string FamilyName { get; set; }
        public DateTime Birthday { get; set; }
        public bool Sex { get; set; }
    }
}
