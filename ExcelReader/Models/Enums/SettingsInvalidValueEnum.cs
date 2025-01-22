using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Enums
{

    /// <summary>
    /// Настройка с возвратом значения если получили ошибку при конвертации
    /// </summary>
    public enum SettingsInvalidValueEnum
    {
        /// <summary>
        /// Возвращает стандартное значение
        /// </summary>
        ReturnDefaultValue = 1,

        /// <summary>
        /// Возвращаем Null
        /// </summary>
        ReturnNull = 2,

        /// <summary>
        /// Пркидываем Exception
        /// </summary>
        ReturnException = 3
    }
}
