using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models;
using ExcelReader.ExcelReader.Models.Enums;
using ExcelReader.ExcelReader.Models.Error.Exceptions;
using ExcelReader.ExcelReader.Models.ReadProperty;
using ExcelReader.ExcelReader.Models.WriteProperty;
using ExcelReader.ExcelReader.Models.WrtieProperty;

namespace ExcelReader.ExcelReader.Reader.SheetHandlers
{

    /// <summary>
    /// Читает лист по заголовку колонок
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    public class ReadSheetByTitleHandler<TModel> : ReadSheetHandler<TModel> where TModel : class, new()
    {
        private int _rowHeaderIndex;

        public ReadSheetByTitleHandler(ISheet workSheet, int rowHeaderIndex) : base(workSheet)
        {
            _rowHeaderIndex = rowHeaderIndex;
        }

        /// <summary>
        /// Сопоставить свойство и колонку по названию в заголовке колонки.
        /// </summary>
        /// <typeparam name="TProperty">Тип свойства</typeparam>
        /// <param name="propertyExpression">Выражение, указывающее на свойство</param>
        /// <param name="columnTitle">Название заголовка</param>
        /// <param name="readExpression">Преобразование значения в колонке</param>
        /// <param name="isReturnException">В случае ошибки преобразование к типу свойста, возвращать дtфолтное значение или выбрасывать ошибку</param>
        /// <returns>Объект для работы с листом</returns>
        /// <exception cref="Exception"></exception>
        public ReadSheetByTitleHandler<TModel> SetColumnByTitle<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression, string columnTitle, Expression<Func<string, object>> readExpression = null, bool isReturnException = false, WritePropertySettings writePropertySettings = null)
        {

            if (propertyExpression.Body is MemberExpression)
            {
                int columnIndex = GetIndexByTitle(columnTitle);

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

            throw new PropertyRequiredException("Вы указали не свойство - " + columnTitle);

        }
        
        /// <summary>
        /// Проверить, что колонка с хаголовкам существует
        /// </summary>
        /// <param name="columnTitle">Значение заголовка</param>
        /// <returns>Индекс колонки</returns>
        /// <exception cref="Exception"></exception>
        private int GetIndexByTitle(string columnTitle)
        {
            var row = _workSheet.GetRow(_rowHeaderIndex);
            if (row == null)
            {
                // Не нашли row
                throw new InvalidRowException("Не нашли строку с заголовками, индекс строки: " + _rowHeaderIndex);
            }

            foreach (var cell in row.Cells)
            {
                if (cell.StringCellValue == columnTitle)
                {
                    return cell.ColumnIndex;
                }
            }

            // Не нашли нужный индекс column по заголовку
            throw new IndexColumnNotFoundException("Не нашли колонку под названием: " + columnTitle);
        }
    }
}
