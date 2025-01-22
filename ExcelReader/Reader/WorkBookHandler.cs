using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models.Enums;

namespace ExcelReader.ExcelReader.Reader
{
    /// <summary>
    /// Класс занимается созданием объекта типа WorkBook для работы с Excel файлом
    /// </summary>
    public class WorkBookHandler
    {

        /// <summary>
        /// Чтение Файла Excel
        /// </summary>
        /// <returns>Объект для работы с файлом Excel</returns>
        /// <param name="path">путь до файла</param>
        /// <exception cref="Exception">ошибка формата</exception>
        public static IWorkbook ReadExcel(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException("Путь к файлу не может быть пустым", nameof(path));
            }
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("Файл не найден", path);
            }
            if (new FileInfo(path).Length == 0)
            {
                throw new InvalidDataException("Файл пустой");
            }
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                IWorkbook workBook = null;
                fileStream.Position = 0;

                string extension = Path.GetExtension(fileStream.Name);

                switch (extension)
                {
                    case ".xls":
                        workBook = new HSSFWorkbook(fileStream);
                        break;
                    case ".xlsx":
                        workBook = new XSSFWorkbook(fileStream);
                        break;
                    default:
                        throw new FormatException("Неизвестный формат");
                }

                return workBook;
            }
        }



        /// <summary>
        /// Создает объект WorkBook
        /// </summary>
        /// <param name="version">Версия excel</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkBook(ExcelVersionEnum? version)
        {
            switch (version)
            {
                case ExcelVersionEnum.Excel07:
                    return new HSSFWorkbook();
                case ExcelVersionEnum.Excel10:
                    return new XSSFWorkbook();
                default:
                    return new HSSFWorkbook();
            }
        }
    }
}
