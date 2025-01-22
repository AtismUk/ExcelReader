using MathNet.Numerics;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models;
using ExcelReader.ExcelReader.Models.Error.Exceptions;
using ExcelReader.ExcelReader.Models.WrtieProperty;
using ExcelReader.Models;

namespace ExcelReader.ExcelReader.Reader.SheetHandlers
{

    /// <summary>
    /// Класс для чтение листа Excle
    /// </summary>
    /// <typeparam name="TModel">Тип модели в которую будут переведенны данные из листа Excel</typeparam>
    public class ReadSheetHandler<TModel> where TModel : class, new()
    {
        protected ISheet _workSheet;
        protected readonly List<WriteProperty<TModel>> _propertyModels = new();

        public ReadSheetHandler(ISheet workSheet)
        {
            _workSheet = workSheet;
        }

        /// <summary>
        /// Читает лист excel и переводит его в список моделей типа <typeparamref name="TModel"></typeparamref>
        /// </summary>
        /// <param name="startRow">С какой строки начинать чтение</param>
        /// <param name="countRows">Количество чтения</param>
        /// <returns>Лист объектов</returns>
        public ResultModel<TModel> ReadSheet(int startRow, int countRows = -1)
        {
            CheckDublicateProperty();

            ResultModel<TModel> result = new();
            int lastRow = CheckRowRange(startRow, countRows);

            for (int row = startRow; row <= lastRow; row++)
            {
                var model = CreateModel(_workSheet.GetRow(row));
                result.Models.Add(model);
            }

            return result;
        }

        /// <summary>
        /// Читает лист excel и переводит его в список моделей типа <typeparamref name="TModel"></typeparamref> асинхронно с поддержкой ProcessState
        /// </summary>
        /// <param name="processState">Состояние процесса</param>
        /// <param name="startRow">С какой строки начинать чтение</param>
        /// <param name="countRows">Количество чтения</param>
        /// <returns>Лист объектов</returns>
        public async Task<ResultModel<TModel>> ReadSheetAsync(ProcessState processState, int startRow, int countRows = -1)
        {
            CheckDublicateProperty();

            ResultModel<TModel> result = new();
            int lastRow = CheckRowRange(startRow, countRows);

            int divider = countRows < 10000 ? 1000 : 10000;

            for (int row = startRow; row <= lastRow; row++)
            {
                var model = CreateModel(_workSheet.GetRow(row));
                result.Models.Add(model);

                await UpdateProcessState();


                async Task UpdateProcessState()
                {
                    processState.Progress = (row - startRow + 1) * 100 / (lastRow - startRow);
                    await Task.Yield();
                }
            }

            return result;

            
        }

        /// <summary>
        /// Читает лист excel и сразу возвращает созданную модель типа <typeparamref name="TModel"></typeparamref>
        /// </summary>
        /// <param name="startRow">С какой строки начинать чтение</param>
        /// <param name="countRows">Количество чтения</param>
        /// <returns>Объект типа <typeparamref name="TModel"></typeparamref></returns>
        public IEnumerable<TModel> ReadSheetYield(int startRow, int countRows = -1)
        {
            CheckDublicateProperty();

            int lastRow = CheckRowRange(startRow, countRows);

            for (int row = startRow; row <= lastRow; row++)
            {
                var model = CreateModel(_workSheet.GetRow(row));
                
                yield return model;
            }
        }

        /// <summary>
        /// Создает модель типа <typeparamref name="TModel"></typeparamref>
        /// </summary>
        /// <param name="row">Строка</param>
        /// <returns>Готовая модель</returns>
        private TModel CreateModel(IRow row)
        {

            TModel model = new();

            foreach (var property in _propertyModels)
            {
                var cell = row.GetCell(property.ColumnIndex);

                var convertedValue = ConvertToPropertyType(property, cell);


                property.CompiledWriteExpression.Invoke(model, convertedValue);
            }
            return model;
        }

        /// <summary>
        /// Конвертирует выражение полученное при указании на свойство в функцию, которая устанавливает значение в это свойство
        /// </summary>
        /// <typeparam name="TProperty">Тип свойста</typeparam>
        /// <param name="propertyExpression">выражение, указывающее на свойство ю</param>
        /// <returns>Скомпелированная функция</returns>
        internal Action<TModel, object> ConvertExpressionToAction<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            if (propertyExpression.Body is MemberExpression memberExpression)
            {
                var modelParameter = Expression.Parameter(typeof(TModel), "model");
                var valueParameter = Expression.Parameter(typeof(object), "value");

                Expression propertyAccess = modelParameter;
                var membres = new Stack<MemberExpression>();

                while (memberExpression != null)
                {
                    membres.Push(memberExpression);
                    memberExpression = memberExpression.Expression as MemberExpression;
                }

                while (membres.Count() > 0)
                {
                    var member = membres.Pop();
                    propertyAccess = Expression.Property(propertyAccess, (PropertyInfo)member.Member);
                }

                var assignExpression = Expression.Assign(propertyAccess, Expression.Convert(valueParameter, typeof(TProperty)));

                var writeAssign = Expression.Lambda<Action<TModel, object>>(assignExpression, modelParameter, valueParameter);

                return writeAssign.Compile();
            }
            throw new ArgumentException("Не смогли преобразовать выражение " + propertyExpression.Body + " в функцию присвоения");
        }

        /// <summary>
        /// Конвертация значения в ячейки в нужный тип
        /// </summary>
        /// <param name="propertyModel">Модель записи данных в свойство</param>
        /// <param name="cell">Ячейка</param>
        /// <returns>Сформированное значение для записи в свойство</returns>
        private object ConvertToPropertyType(WriteProperty<TModel> propertyModel, ICell cell)
        {
            object rawValue = GetCellValue(cell);

            if (propertyModel.CompliedConvertExpression != null)
            {
                try
                {
                    return propertyModel.CompliedConvertExpression(rawValue?.ToString());
                }
                catch
                {
                    return GetDefaultValue(propertyModel.PropertyType);
                }
            }

            if (rawValue == null)
            {
                return HandleNullValue(propertyModel);
            }

            if (propertyModel.PropertyType == typeof(string))
            {
                return rawValue.ToString();
            }

            Type targetType = propertyModel.Nullable
                            ?? propertyModel.PropertyType;

            try
            {
                if (targetType == typeof(DateTime))
                {
                    return ConvertToDateTime(rawValue);
                }

                return Convert.ChangeType(rawValue, targetType, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                if (propertyModel.IsReturnException)
                {
                    throw new InvalidCastException(
                        $"Ошибка конвертации ячейки {cell?.Address} в свойство " +
                        $"{GetPropertyName(propertyModel)}: {ex.Message}", ex);
                }

                return GetDefaultValue(targetType);
            }
        }
        /// <summary>
        /// Проверяет диапозон строк на валидность
        /// </summary>
        /// <param name="startRow">С какой строки начинать чтение</param>
        /// <param name="countRows">Сколько строк читать</param>
        /// <returns>Индекс строки на какой заканчивать чтение</returns>
        /// <exception cref="Exception"></exception>
        private int CheckRowRange(int startRow = 0, int countRows = -1)
        {
            IRow row = _workSheet.GetRow(startRow);
            if (row == null)
            {
                throw new InvalidExcelRowRangeException($"В листе не существует строки с индексом, указанный в startRow: {startRow}");
            }

            if (startRow >= _workSheet.LastRowNum)
            {
                throw new InvalidExcelRowRangeException($"Количество строк в Excel листе меньше значения, от куда вы указали читать: {startRow} >= {_workSheet.LastRowNum}");
            }

            int lastRow = countRows == -1 ? _workSheet.LastRowNum : countRows;
            if (lastRow == 0)
            {
                lastRow = countRows + startRow;

                if (lastRow > _workSheet.LastRowNum)
                {
                    throw new InvalidExcelRowRangeException($"Количество строк в Excel файле меньше, чем вы указали до куда читать лист: {lastRow} >= {_workSheet.LastRowNum}");
                }

                if (lastRow <= 0)
                {
                    throw new InvalidExcelRowRangeException($"Вы указали количество строк на чтение как 0 или отрицательное число {countRows}");
                }
            }

            return lastRow;
        }

        /// <summary>
        /// Проверка на дурока, проверяет не указал ли пользователь свойство для записи несколько раз
        /// </summary>
        /// <exception cref="DuplicatePropertyException">Указал свойство несколько раз</exception>
        private void CheckDublicateProperty()
        {
            int y = 1;
            foreach (var property in _propertyModels)
            {
                for (int i = y; i < _propertyModels.Count; i++)
                {
                    if (property.ExpressionBody.ToString() == _propertyModels[i].ExpressionBody.ToString())
                    {
                        throw new DuplicatePropertyException("Вы указали одно и тоже свойство несколько раз: " + property.ExpressionBody);
                    }
                }
                y++;
              
            }
        }



        /// <summary>
        /// Получения значения ячейки
        /// </summary>
        private static object GetCellValue(ICell cell)
        {
            if (cell == null) return null;

            return cell.CellType switch
            {
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                    ? cell.DateCellValue
                    : cell.NumericCellValue,
                CellType.Boolean => cell.BooleanCellValue,
                CellType.Formula => GetFormulaValue(cell),
                CellType.Error => cell.ErrorCellValue,
                CellType.Blank => null,
                _ => cell.ToString()
            };
        }

        /// <summary>
        /// Получения значения ячейки из формулы
        /// </summary>
        private static object GetFormulaValue(ICell cell)
        {
            try
            {
                return cell.CachedFormulaResultType switch
                {
                    CellType.Numeric => cell.NumericCellValue,
                    CellType.String => cell.StringCellValue,
                    CellType.Boolean => cell.BooleanCellValue,
                    _ => cell.ToString()
                };
            }
            catch
            {
                return cell.ToString();
            }
        }

        /// <summary>
        /// Конвертирует значение в DateTime
        /// </summary>
        private static DateTime ConvertToDateTime(object rawValue)
        {
            if (rawValue is DateTime dt)
                return dt;

            if (rawValue is double d)
                return DateTime.FromOADate(d);

            if (rawValue is string s)
            {
                var culture = CultureInfo.CreateSpecificCulture("ru-RU"); 
                return DateTime.Parse(s, culture);
            }

            throw new InvalidCastException($"Не удалось преобразовать {rawValue} в DateTime");
        }

        /// <summary>
        /// Обрабатывает null-значения
        /// </summary>
        private static object HandleNullValue(WriteProperty<TModel> propertyModel)
        {
            return propertyModel.Nullable == null ? Activator.CreateInstance(propertyModel.PropertyType) : null;
        }

        /// <summary>
        /// Возвращает значение по умолчанию для типа
        /// </summary>
        private static object GetDefaultValue(Type type)
        {
            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }

        /// <summary>
        /// Возвращает имя свойства
        /// </summary>
        private string GetPropertyName(WriteProperty<TModel> propertyModel)
        {
            return (propertyModel.ExpressionBody as MemberExpression)?.Member.Name ?? "Unknown";
        }
    }
}
