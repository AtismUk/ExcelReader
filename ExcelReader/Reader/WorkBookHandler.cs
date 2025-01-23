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
        /// Прочтать Excel через stream
        /// </summary>
        /// <param name="stream">стрим</param>
        /// <returns>Объект для работы с файлом Excel</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="InvalidDataException"></exception>
        /// <exception cref="FormatException"></exception>
        public static IWorkbook ReadStream(Stream stream)
        {

            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            if (!stream.CanRead)
                throw new ArgumentException("Поток не доступен для чтения", nameof(stream));


            byte[] header = new byte[8];
            int bytesRead = stream.Read(header, 0, header.Length);
            stream.Seek(-bytesRead, SeekOrigin.Current);

            if (bytesRead < 8)
                throw new InvalidDataException("Недостаточно данных для определения формата");

            try
            {
                if (header[0] == 0x50 && header[1] == 0x4B && header[2] == 0x03 && header[3] == 0x04)
                    return new XSSFWorkbook(stream);

                if (header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0
                    && header[4] == 0xA1 && header[5] == 0xB1 && header[6] == 0x1A && header[7] == 0xE1)
                    return new HSSFWorkbook(stream);

            }
            catch (Exception ex)
            {
                throw new FormatException("Формат не соответсмтвует xlx или xlsx");
            }

            throw new FormatException("Формат не соответсмтвует xlx или xlsx");

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
