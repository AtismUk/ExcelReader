using ExcelReader.ExcelReader.Reader.SheetHandlers;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models;
using ExcelReader.ExcelReader.Models.Enums;
using ExcelReader.ExcelReader.Models.ReadProperty;


namespace ExcelReader.ExcelReader.Reader
{

    /// <summary>
    /// Создание Excel файла
    /// </summary>
    public class WriteExcelHandler
    {
        private IWorkbook _workBook;

        List<WriteSheetHandler<object>> SheetHandlers;

        public WriteExcelHandler(IWorkbook workBook)
        {
            _workBook = workBook;
        }


        /// <summary>
        /// Создание листа
        /// </summary>
        /// <typeparam name="TModel">Тип модели,которая будет записываться в лист</typeparam>
        /// <param name="sheetName">Название листа</param>
        /// <param name="models">Лист моделей</param>
        /// <param name="commonHeaderStyle">Общий стиль для заголовка</param>
        /// <param name="commonDataStyle">Общий стьиль для данных</param>
        /// <returns></returns>
        public WriteSheetHandler<TModel> CreateSheet<TModel>(string sheetName, List<TModel> models, ICellStyle commonHeaderStyle = null, ICellStyle commonDataStyle = null) 
        {
            WriteSheetHandler<TModel> writeSheetHandler = new(_workBook, sheetName, models, commonHeaderStyle: commonHeaderStyle, commonDataStyle: commonDataStyle);
            
            return writeSheetHandler;
        }
        public WriteSheetHandler<TModel> CreateSheet<TModel>(string sheetName, TModel model, ICellStyle commonHeaderStyle = null, ICellStyle commonDataStyle = null)
        {
            List<TModel> models = new();
            models.Add(model);
            WriteSheetHandler<TModel> writeSheetHandler = new(_workBook, sheetName, models, commonHeaderStyle: commonHeaderStyle, commonDataStyle: commonDataStyle);

            return writeSheetHandler;
        }


        /// <summary>
        /// Создать файл
        /// </summary>
        /// <param name="fileInfo">Настройки файла</param>
        public void CreateExcelFile(ExcelFileInfo fileInfo)
        {
            string extension = "";

            if (_workBook is HSSFWorkbook)
            {
                extension = ".xls";
            }
            else if (_workBook is XSSFWorkbook)
            {
                extension = ".xlsx";
            }

            using (FileStream fileStream = new FileStream(fileInfo.Path + fileInfo.Name + extension, FileMode.Create))
            {
                _workBook.Write(fileStream, false);
            }
        }


        /// <summary>
        /// Записать Excel в стрим
        /// </summary>
        /// <param name="stream">Стрим</param>
        /// <param name="leaveOpen">Оставить стрим открытым для записи после завершения метода</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public void WriteToStream(Stream stream, bool leaveOpen = false)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            if (!stream.CanWrite)
                throw new ArgumentException("Поток недоступен для записи", nameof(stream));

            long originalPosition = stream.CanSeek ? stream.Position : 0;

            try
            {
                _workBook.Write(stream, leaveOpen);

                if (stream.CanSeek)
                    stream.Seek(0, SeekOrigin.Begin);

            }
            finally
            {
                if (stream.CanSeek)
                    stream.Position = originalPosition;

                if (!leaveOpen)
                {
                    stream.Close();
                }
            }
        }
    }
}
