using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models.WrtieProperty;

namespace ExcelReader.ExcelReader.Models.ReadProperty
{
    public class ReadProperty<TModel>
    {

        public LambdaExpression PropertyExpression { get; }

        public Func<TModel, string> CompliedReadExpression { get; set; }

        public string ColumnName { get; set; }

        public ICellStyle HeaderCellStyle { get; set; }
        public ICellStyle DataCellStyle { get; set; }


    }
}
