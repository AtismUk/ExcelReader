using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.ExcelReader.Models.Enums
{
    public enum CommonStateEnum
    {
        /// <summary>
        /// Неизвестный тип
        /// </summary>
        None,
        /// <summary>
        /// В ожидании
        /// </summary>
        Wait,
        /// <summary>
        /// Выполняется
        /// </summary>
        InProgress,
        /// <summary>
        /// Выполнено успешно
        /// </summary>
        Success,
        /// <summary>
        /// Ошибка при выполнении
        /// </summary>
        Failed
    }
}
