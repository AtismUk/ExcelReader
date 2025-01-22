using ExcelReader.ExcelReader.Models.Enums;
using ExcelReader.ExcelReader.Reader;
using ExcelReader.Models;
using TestExcelReader.Models;

namespace TestExcelReader
{
    internal class Program
    {
        static List<TestModel> models = new();

        static List<TestModel> sheetModels1 = new();
        static List<TestModel> sheetModels2 = new();

        static ProcessState ProcessStateList1 = new();
        static ProcessState ProcessStateList2 = new();
        static ProcessState WriteProcessStateList1 = new();
        static ProcessState WriteProcessStateList2 = new();

        static async Task Main(string[] args)
        {

            ReadExcelIndex();

            WriteExcel();

            Console.WriteLine("Закончил");

        }

        static void ReadExcelIndex()
        {
            using (var woorkBook = WorkBookHandler.ReadExcel(Path.Combine(Environment.CurrentDirectory, "Test.xls")))
            {
                var sheetFactory = new ReadSheetFactory(woorkBook);

                var sheetHandler1 = sheetFactory.SetSheetByIndex<TestModel>(indexSheet: 0)
                    .SetColumnByIndex(x => x.Id, 0)
                    .SetColumnByIndex(x => x.Person.Name, 1)
                    .SetColumnByIndex(x => x.Person.Surname, 2)
                    .SetColumnByIndex(x => x.Person.FamilyName, 3, isReturnException: true)
                    .SetColumnByIndex(x => x.Person.Birthday, 4)
                    .SetColumnByIndex(x => x.Person.Sex, 5, x => x == "М" ? true : false)
                    .SetColumnByIndex(x => x.Document.Seria, 6)
                    .SetColumnByIndex(x => x.Document.Number, 7)
                    .SetColumnByIndex(x => x.TestNull, 8);

                models = sheetHandler1.ReadSheet(startRow: 1).Models;
            }
        }

        static void ReadExcelTitle()
        {
            using (var woorkBook = WorkBookHandler.ReadExcel(Path.Combine(Environment.CurrentDirectory, "Test.xls")))
            {
                var sheetFactory = new ReadSheetFactory(woorkBook);

                var sheetHandler1 = sheetFactory.SetSheetByIndex<TestModel>(indexSheet: 0, rowHeaderIndex: 0)
                    .SetColumnByTitle(x => x.Id, "Id")
                    .SetColumnByTitle(x => x.Person.Name, "Name")
                    .SetColumnByTitle(x => x.Person.Surname, "Surname")
                    .SetColumnByTitle(x => x.Person.FamilyName, "FamilyName")
                    .SetColumnByTitle(x => x.Person.Birthday, "Birthday")
                    .SetColumnByTitle(x => x.Person.Sex, "Sex", x => x == "М" ? true : false)
                    .SetColumnByTitle(x => x.Document.Seria, "Seria")
                    .SetColumnByTitle(x => x.Document.Number, "Number");

                models = sheetHandler1.ReadSheet(startRow: 1, countRows: 2000).Models;
            }
        }

        static void WriteExcel()
        {
            using (var workBook = WorkBookHandler.CreateWorkBook(ExcelVersionEnum.Excel07))
            {

                var writeExcelHandler = new WriteExcelHandler(workBook);

                writeExcelHandler.CreateSheet(sheetName: "Лист тест 1", models)
                    .SetColumn(x => x.Id, "Id")
                    .SetColumn(x => x.Person.Name, "Name")
                    .SetColumn(x => x.Person.Surname, "Surname")
                    .SetColumn(x => x.Person.FamilyName, "FamilyName")
                    .SetColumn(x => x.Person.Birthday.ToShortDateString(), "Birthday")
                    .SetColumn(x => x.Person.Sex == true ? "Male" : "Female", "Sex")
                    .SetColumn(x => x.Document.Seria.ToString(), "Seria")
                    .SetColumn(x => x.Document.Number.ToString(), "Number")
                    .SetColumn(x =>
                    {
                        if (x.Document.Number != 0 || x.Document.Seria != 0)
                        {
                            return "Д";
                        }
                        return "Н";
                    }, "DLO")
                    .WriteModels();

                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                string projectPath = Path.GetFullPath(Path.Combine(basePath, @"..\..\.."));                string excelFolderPath = Path.Combine(projectPath, @"Excel\");

                if (!Directory.Exists(excelFolderPath))
                {
                    Directory.CreateDirectory(excelFolderPath);
                }

                string FileName = $"Test {Path.GetRandomFileName()}";

                writeExcelHandler.CreateExcelFile(new(FileName, excelFolderPath));
            }
        }


        #region Тест асинхронности

        static async Task StartAsyncTest()
        {
            var stateTask = StateAsync();


            await ReadExcelAsync();

            // я не знаю почему поток блокируется при записи и прогресс не отображается :(

            await WriteExcelAsync();


            await stateTask;
        }

        static async Task StateAsync()
        {
            while (ProcessStateList1.Progress != 100 && WriteProcessStateList2.Progress != 100)
            {
                Console.Clear();
                Console.WriteLine("Чтение листа 1: " + ProcessStateList1.Progress + "%");
                Console.WriteLine("Чтение листа 2: " + ProcessStateList2.Progress + "%");
                Console.WriteLine("------------------------------------------------------");
                Console.WriteLine("Запись в Лист 1: " + WriteProcessStateList1.Progress + "%");
                Console.WriteLine("Запись в Лист 2: " + WriteProcessStateList2.Progress + "%");
                await Task.Delay(50);
            }
        }

        static async Task WriteExcelAsync()
        {
            using (var workBook = WorkBookHandler.CreateWorkBook(ExcelVersionEnum.Excel07))
            {
                var writeExcelHandler = new WriteExcelHandler(workBook);

                await writeExcelHandler.CreateSheet("Лист тест 1", sheetModels1)
                   .SetColumn(x => x.Id, "Id")
                   .SetColumn(x => x.Person.Name, "Name")
                   .SetColumn(x => x.Person.Surname, "Surname")
                   .SetColumn(x => x.Person.FamilyName, "FamilyName")
                   .SetColumn(x => x.Person.Birthday.ToShortDateString(), "Birthday")
                   .SetColumn(x => x.Person.Sex == true ? "Male" : "Female", "Sex")
                   .SetColumn(x => x.Document.Seria.ToString(), "Seria")
                   .SetColumn(x => x.Document.Number.ToString(), "Number")
                   .WriteModelsAsync(WriteProcessStateList1);



                await writeExcelHandler.CreateSheet("Лист тест 2", sheetModels2)
                     .SetColumn(x => x.Id, "Id")
                     .SetColumn(x => x.Person.Name, "Name")
                     .SetColumn(x => x.Person.Surname, "Surname")
                     .SetColumn(x => x.Person.FamilyName, "FamilyName")
                     .SetColumn(x => x.Person.Birthday.ToShortDateString(), "Birthday")
                     .SetColumn(x => x.Person.Sex == true ? "Male" : "Female", "Sex")
                     .SetColumn(x => x.Document.Seria.ToString(), "Seria")                     .SetColumn(x => x.Document.Number.ToString(), "Number")
                     .WriteModelsAsync(WriteProcessStateList2);

                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                string projectPath = Path.GetFullPath(Path.Combine(basePath, @"..\..\.."));
                string excelFolderPath = Path.Combine(projectPath, @"Excel\");

                if (!Directory.Exists(excelFolderPath))
                {
                    Directory.CreateDirectory(excelFolderPath);
                }

                string FileName = $"Test Async {Path.GetRandomFileName()}";

                writeExcelHandler.CreateExcelFile(new(FileName, excelFolderPath));

            }
        }

        static async Task ReadExcelAsync()
        {
            using (var woorkBook = WorkBookHandler.ReadExcel(Path.Combine(Environment.CurrentDirectory, "Test.xls")))
            {
                var sheetFactory = new ReadSheetFactory(woorkBook);

                var sheetHandler1 = sheetFactory.SetSheetByIndex<TestModel>(indexSheet: 0)
                    .SetColumnByIndex(x => x.Id, 0)
                    .SetColumnByIndex(x => x.Person.Name, 1)
                    .SetColumnByIndex(x => x.Person.Surname, 2)
                    .SetColumnByIndex(x => x.Person.FamilyName, 3, isReturnException: true)
                    .SetColumnByIndex(x => x.Person.Birthday, 4)
                    .SetColumnByIndex(x => x.Person.Sex, 5, x => x == "М" ? true : false)
                    .SetColumnByIndex(x => x.Document.Seria, 6)
                    .SetColumnByIndex(x => x.Document.Number, 7);



                var sheetHandler2 = sheetFactory.SetSheetByIndex<TestModel>(indexSheet: 0)
                    .SetColumnByIndex(x => x.Id, 0)
                    .SetColumnByIndex(x => x.Person.Name, 1)
                    .SetColumnByIndex(x => x.Person.Surname, 2)
                    .SetColumnByIndex(x => x.Person.FamilyName, 3)
                    .SetColumnByIndex(x => x.Person.Birthday, 4)
                    .SetColumnByIndex(x => x.Person.Sex, 5, x => x == "М" ? true : false)
                    .SetColumnByIndex(x => x.Document.Seria, 6)
                    .SetColumnByIndex(x => x.Document.Number, 7);

                var readFirstSheet = sheetHandler1.ReadSheetAsync(ProcessStateList1, 1);
                var readSecondSheet = sheetHandler2.ReadSheetAsync(ProcessStateList2, 1);

                var results = await Task.WhenAll(readFirstSheet, readSecondSheet);

                sheetModels1 = results[0].Models;
                sheetModels2 = results[1].Models;
            }
        }

        #endregion


        #region Тесты

        #endregion
    }
}
