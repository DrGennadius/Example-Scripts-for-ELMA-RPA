using System;
using System.IO;
using System.Collections.Generic;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Контекст процесса
    /// <summary>
    public class Context
    {
        /// <summary>
        /// Рабочая папка, где сохранять вложения
        /// </summary>
        public String WorkDirectory
        {
            get;
            set;
        }

        /// <summary>
        /// Контексты писем как строки
        /// </summary>
        public List<String> MessageContextStrings
        {
            get;
            set;
        }

        /// <summary>
        /// Хост
        /// </summary>
        public String Host
        {
            get;
            set;
        }

        /// <summary>
        /// Порт
        /// </summary>
        public Nullable<Int64> Port
        {
            get;
            set;
        }

        /// <summary>
        /// Использовать SSL
        /// </summary>
        public Nullable<Boolean> UseSsl
        {
            get;
            set;
        }

        /// <summary>
        /// Имя пользователя
        /// </summary>
        public String UserName
        {
            get;
            set;
        }

        /// <summary>
        /// Пароль
        /// </summary>
        public String Password
        {
            get;
            set;
        }

        /// <summary>
        /// Тема. Текст для поиска в тексте темы
        /// </summary>
        public String SubjectText
        {
            get;
            set;
        }

        /// <summary>
        /// Отправитель. Текст для поиска в тексте отправителя
        /// </summary>
        public String SenderText
        {
            get;
            set;
        }
    }
}
