using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models.Error.Exceptions;
using ExcelReader.ExcelReader.Reader.SheetHandlers;

namespace ExcelReader.ExcelReader.Reader
{

    /// <summary>
    /// Фабрика по созданию нужной реализации для работы с чтеним листа Excel
    /// </summary>
    public class ReadSheetFactory
    {
        private IWorkbook _workBook;
        public ReadSheetFactory(IWorkbook workBook)
        {
            _workBook = workBook;
        }

        /// <summary>
        /// Установить лист для чтения в документе
        /// </summary>
        /// <param name="indexSheet">Индекс листа</param>
        /// <typeparam name="TModel">модель, в которую будет переведен лист</typeparam>
        /// <returns>объект для работы с листом, который сопоставляет колонки со свойством по индексу колонки</returns>
        public ReadSheetByIndexHandler<TReadModel> SetSheetByIndex<TReadModel>(int indexSheet) where TReadModel : class, new()
        {
            var Sheet = CheckSheetIndex(indexSheet);
            return new ReadSheetByIndexHandler<TReadModel>(Sheet);
        }

        /// <summary>
        /// Установить лист для чтения в документе
        /// </summary>
        /// <param name="sheetName">имя листа</param>
        /// <typeparam name="TModel">модель, в которую будет переведен лист</typeparam>
        /// <returns>объект для работы с листом, который сопоставляет колонки со свойством по индексу колонки</returns>
        public ReadSheetByIndexHandler<TReadModel> SetSheetByName<TReadModel>(string sheetName) where TReadModel : class, new()
        {
            var Sheet = CheckSheetName(sheetName);
            return new ReadSheetByIndexHandler<TReadModel>(Sheet);
        }

        /// <summary>
        /// Установить лист для чтения в документе
        /// </summary>
        /// <param name="indexSheet">индекс листа</param>
        /// <param name="startRowHeaderIndex">строка с заголовками</param>
        /// <typeparam name="TModel">модель, в которую будет переведен лист</typeparam>
        /// <returns>объект для работы с листом, который сопоставляет колонки со свойством по имени колонки в заголовке</returns>
        public ReadSheetByTitleHandler<TReadModel> SetSheetByIndex<TReadModel>(int indexSheet, int rowHeaderIndex) where TReadModel : class, new()
        {
            var Sheet = CheckSheetIndex(indexSheet);
            return new ReadSheetByTitleHandler<TReadModel>(Sheet, rowHeaderIndex);
        }

        /// <summary>
        /// Установить лист для чтения в документе
        /// </summary>
        /// <param name="sheetName">имя листа</param>
        /// <param name="startRowHeaderIndex">строка с заголовками</param>
        /// <typeparam name="TModel">модель, в которую будет переведен лист</typeparam>
        /// <returns>объект для работы с листом, который сопоставляет колонки со свойством по имени колонки в заголовке</returns>
        public ReadSheetByTitleHandler<TReadModel> SetSheetByName<TReadModel>(string sheetName, int rowHeaderIndex) where TReadModel : class, new()
        {
            var Sheet = CheckSheetName(sheetName);
            return new ReadSheetByTitleHandler<TReadModel>(Sheet, rowHeaderIndex);
        }

        private ISheet CheckSheetName(string sheetName)
        {
            ISheet Sheet = _workBook.GetSheet(sheetName);
            if (Sheet == null)
            {
                throw new IndexSheetNotFoundException($"Лист под именем {sheetName} не найден в документе");
            }
            return Sheet;
        }

        private ISheet CheckSheetIndex(int sheetIndex)
        {
            try
            {
                ISheet Sheet = _workBook.GetSheetAt(sheetIndex);
                return Sheet;
            }
            catch (Exception)
            {
                throw new IndexSheetNotFoundException($"Лист под индексом {sheetIndex} не найден в документе");
            }
        }
    }
}
