using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models
{
    public class ResultModel<TModel>
    {
        public List<TModel> Models { get; set; } = new List<TModel>();
    }
}
