using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models;
using ExcelReader.ExcelReader.Models.ReadProperty;
using ExcelReader.ExcelReader.Reader;
using NPOI.SS.Formula.Functions;
using ExcelReader.Models;

namespace ExcelReader.ExcelReader.Reader.SheetHandlers
{

    /// <summary>
    /// Запись данных в лист
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    public class WriteSheetHandler<TModel>
    {
        private IWorkbook _workBook;
        private List<ReadProperty<TModel>> _propertyModels = new();
        private List<TModel> _writeModels;
        private ICellStyle _commonHeaderStyle;
        private ICellStyle _commonDataStyle;
        private string _sheetName;


        public WriteSheetHandler(IWorkbook workBook, string sheetName, List<TModel> models, ICellStyle commonHeaderStyle = null, ICellStyle commonDataStyle = null)
        {
            _workBook = workBook;
            if (models.Count() == 0)
            {
                throw new Exception("Список моделей не может быть пустым");
            }
            _writeModels = models;

            _commonHeaderStyle = commonHeaderStyle == null ? _commonHeaderStyle = workBook.CreateCellStyle() : commonHeaderStyle;
            _commonDataStyle = commonDataStyle == null ? _commonDataStyle = workBook.CreateCellStyle() : commonDataStyle;
            _sheetName = sheetName;
        }

        /// <summary>
        /// Сопоставить свойство с колонкой куда будут записываться данные
        /// </summary>
        /// <param name="propertyExpression">Получение значения</param>
        /// <param name="columnName">Название колонки</param>
        /// <param name="headerStyle">Стиль заголовка</param>
        /// <param name="dataStyle">Стиль данных</param>
        /// <returns>Объект для записи данных в Excel</returns>
        public WriteSheetHandler<TModel> SetColumn(Func<TModel, string> propertyExpression, string columnName, ICellStyle headerStyle = null, ICellStyle dataStyle = null)
        {
            var propertyModel = new ReadProperty<TModel>()
            {
                CompliedReadExpression = x => propertyExpression(x),
                ColumnName = columnName,
                HeaderCellStyle = headerStyle == null ? _commonHeaderStyle : headerStyle,
                DataCellStyle = dataStyle == null ? _commonDataStyle : dataStyle
            };

            _propertyModels.Add(propertyModel);

            return this;
        }

        /// <summary>
        /// Сопоставить свойство с колонкой куда будут записываться данные
        /// </summary>
        /// <param name="func">Делегат получения значения</param>
        /// <param name="columnName">Название колонки</param>
        /// <param name="headerStyle">Стиль заголовка</param>
        /// <param name="dataStyle">Стиль данных</param>
        /// <returns>Объект для записи данных в Excel</returns>
        public WriteSheetHandler<TModel> SetColumnFunc(Func<string> func, string columnName, ICellStyle headerStyle = null, ICellStyle dataStyle = null)
        {
            return SetColumn(x => func(), columnName, headerStyle, dataStyle);
        }

        /// <summary>
        /// Создать лист и записать в него данные
        /// </summary>
        /// <param name="sheetName">Название листа</param>
        public void WriteModels()
        {
            var sheet = _workBook.CreateSheet(_sheetName);

            WriteHeaders(sheet);

            int line = 1;

            foreach (var model in _writeModels)
            {
                var row = sheet.CreateRow(line);
                for (int y = 0; y < _propertyModels.Count(); y++)
                {
                    var value = _propertyModels[y].CompliedReadExpression(model);

                    ICell cell = row.CreateCell(y);

                    if (_propertyModels[y].DataCellStyle != null)
                    {
                        cell.CellStyle = _propertyModels[y].DataCellStyle;
                    }

                    cell.SetCellValue(value);
                }

                line++;
            }

            _propertyModels.Clear();
        }

        /// <summary>
        /// Создать лист и записать в него данные асинхронно с поддержкой ProcessState
        /// </summary>
        /// <param name="sheetName">Название листа</param>
        /// <param name="processState">Процесс состояния</param>
        /// <returns></returns>
        public async Task WriteModelsAsync(ProcessState processState)
        {
            var sheet = _workBook.CreateSheet(_sheetName);

            WriteHeaders(sheet);

            int line = 1;
            int divider = _writeModels.Count < 10000 ? 1000 : 10000;


            foreach (var model in _writeModels)
            {
                processState.Progress = line * 100 / _writeModels.Count;


                await Parallel.ForEachAsync(_writeModels, async (model, _) =>
                {
                    var row = sheet.CreateRow(line);
                    for (int y = 0; y < _propertyModels.Count; y++)
                    {
                        var value = _propertyModels[y].CompliedReadExpression(model);
                        var cell = row.CreateCell(y);
                        cell.SetCellValue(value);
                    }

                    if (line % divider == 0 && line == _writeModels.Count)
                    {
                        await Task.Delay(1);
                    }
                });

                line++;
            }

            _propertyModels.Clear();

        }

        /// <summary>
        /// Записать заколовок
        /// </summary>
        /// <param name="workSheet"></param>
        private void WriteHeaders(ISheet workSheet)
        {
            var row = workSheet.CreateRow(0);

            for (int i = 0; i < _propertyModels.Count(); i++)
            {
                var cell = row.CreateCell(i);

                if (_propertyModels[i].HeaderCellStyle != null)
                {
                    cell.CellStyle = _propertyModels[i].HeaderCellStyle;
                }

                cell.SetCellValue(_propertyModels[i].ColumnName);
            }
        }

    }
}
