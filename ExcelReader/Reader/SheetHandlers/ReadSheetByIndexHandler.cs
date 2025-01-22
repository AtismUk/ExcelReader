using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models;
using ExcelReader.ExcelReader.Models.Enums;
using ExcelReader.ExcelReader.Models.Error.Exceptions;
using ExcelReader.ExcelReader.Models.WriteProperty;
using ExcelReader.ExcelReader.Models.WrtieProperty;

namespace ExcelReader.ExcelReader.Reader.SheetHandlers
{

    /// <summary>
    /// Читает лист по индексу колонок
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    public class ReadSheetByIndexHandler<TModel> : ReadSheetHandler<TModel> where TModel : class, new()
    {

        public ReadSheetByIndexHandler(ISheet workSheet) : base(workSheet)
        {

        }

        /// <summary>
        /// Сопоставить свойство и колонку по индексу колонки
        /// </summary>
        /// <typeparam name="TProperty">Тип свойства</typeparam>
        /// <param name="propertyExpression">Выражение, указывающее на свойство</param>
        /// <param name="columnIndex">Индекс колонки</param>
        /// <param name="readExpression">Преобразование значения в колонке</param>
        /// <param name="isReturnException">В случае ошибки преобразование к типу свойста, возвращать дефолтное значение или выбрасывать ошибку</param>
        /// <returns>Объект для работы с листом</returns>
        /// <exception cref="Exception"></exception>
        public ReadSheetByIndexHandler<TModel> SetColumnByIndex<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression, int columnIndex, Expression<Func<string, object>> readExpression = null, bool isReturnException = false, WritePropertySettings writePropertySettings = null)
        {

            if (propertyExpression.Body is MemberExpression)
            {
                var propertyModel = new WriteProperty<TModel>()
                {
                    ColumnIndex = columnIndex,
                    CompiledWriteExpression = ConvertExpressionToAction(propertyExpression),
                    PropertyType = typeof(TProperty),
                    WritePropertySettings = writePropertySettings,
                    ExpressionBody = propertyExpression.Body,
                    IsReturnException = isReturnException,
                    Nullable = Nullable.GetUnderlyingType(typeof(TProperty))
                };

                if (readExpression != null)
                {
                    propertyModel.CompliedConvertExpression = readExpression.Compile();
                }

                _propertyModels.Add(propertyModel);

                return this;
            }

            throw new PropertyRequiredException("Вы указали не свойство - " + columnIndex);
        }

    }
}
