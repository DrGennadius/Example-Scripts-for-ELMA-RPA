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
        /// Текст сообщения
        /// </summary>
        public String MessageText
        {
            get;
            set;
        }

        /// <summary>
        /// Вложения как FileInfo
        /// </summary>
        public List<FileInfo> Attachments
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
        /// Тема
        /// </summary>
        public String Subject
        {
            get;
            set;
        }

        /// <summary>
        /// Отправитель
        /// </summary>
        public String Sender
        {
            get;
            set;
        }

        /// <summary>
        /// Получатель
        /// </summary>
        public String Reciver
        {
            get;
            set;
        }
    }
}