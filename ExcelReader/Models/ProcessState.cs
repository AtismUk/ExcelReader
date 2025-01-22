using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.ExcelReader.Models.Enums;

namespace ExcelReader.Models
{
    public class ProcessState
    {
        private CommonStateEnum state = CommonStateEnum.None;
        private int progress = 0;
        private string message;

        private int currentAction = 0;

        /// <summary>
        /// Ключ состояния процесса
        /// </summary>
        public Guid Key { get; private set; }
        /// <summary>
        /// Базовое состояние процесса
        /// </summary>
        public CommonStateEnum State
        {
            get => state;
            set
            {
                state = value;
                LastModified = DateTime.Now;
            }
        }
        /// <summary>
        /// Прогресс выполнения процесса в процентах
        /// </summary>
        /// <remarks>Так же может быть изменено с помощью свойств <see cref="CurrentAction" /> и <see cref="TotalActions" /></remarks>
        public int Progress
        {
            get => progress;
            set
            {
                if (value >= 0 && value <= 100)
                {
                    progress = value;
                    // Если вручную задаём прогресс, то сбрасываем счётчики действий
                    currentAction = 0; TotalActions = 0;
                    LastModified = DateTime.Now;
                }
            }
        }
        /// <summary>
        /// Сообщение выполнения процесса
        /// </summary>
        public string Message
        {
            get => message;
            set
            {
                message = value;
                LastModified = DateTime.Now;
            }
        }
        /// <summary>
        /// Временная отметка последнего изменения состояния процесса
        /// </summary>
        public DateTime LastModified { get; private set; } = DateTime.Now;

        /// <summary>
        /// Общее количество действий, которые планируется выполнить
        /// </summary>
        public int TotalActions { get; set; }
        /// <summary>
        /// Номер текущего выполняемого действия
        /// </summary>
        public int CurrentAction
        {
            get => currentAction;
            set
            {
                currentAction = value;
                //Считаем прогресс самостоятельно
                progress = (int)((double)currentAction / TotalActions * 100);
                if (progress > 100)
                    progress = 100;
                LastModified = DateTime.Now;
            }
        }

        public ProcessState()
        {
            Key = Guid.NewGuid();
        }

        public ProcessState(Guid key)
        {
            Key = key;
        }
    }

}
