using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models.Enums;
using ExcelReader.ExcelReader.Models.WriteProperty;

namespace ExcelReader.ExcelReader.Models.WrtieProperty
{
    public class WriteProperty<TModel>
    {
        public Action<TModel, object> CompiledWriteExpression { get; set; }
        public Func<string, object> CompliedConvertExpression { get; set; }
        public int ColumnIndex { get; set; }
        public Type PropertyType { get; set; }
        public WritePropertySettings WritePropertySettings { get; set; }
        public Expression ExpressionBody { get; set; }
        public bool IsReturnException { get; set; } = false;
        public Type? Nullable { get; set; }

    }
}
