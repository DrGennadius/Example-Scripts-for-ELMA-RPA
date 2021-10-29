using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Контекст письма (тело сообщения + пути к вложениям).
    /// </summary>
    public class MessageContext
    {
        public MessageContext()
        {
            AttachmentPaths = new List<string>();
        }

        public MessageContext(string contextString)
            : this()
        {
            SetContext(contextString);
        }

        /// <summary>
        /// Тело письма
        /// </summary>
        public string TextBody { get; set; }

        /// <summary>
        /// Список путей загруженных файлов вложений
        /// </summary>
        public List<string> AttachmentPaths { get; set; }

        /// <summary>
        /// Задать контекст
        /// </summary>
        /// <param name="contextString"></param>
        public void SetContext(string contextString)
        {
            var options = new JsonSerializerOptions
            {
                // Кодировка для Unicode: Basic Latin и Cyrillic
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
            };
            // Десериализация
            var context = JsonSerializer.Deserialize<MessageContext>(contextString, options);
            TextBody = context.TextBody;
            AttachmentPaths = context.AttachmentPaths;
            if (AttachmentPaths == null)
            {
                AttachmentPaths = new List<string>();
            }
        }

        public string GetContextJSON()
        {
            var options = new JsonSerializerOptions
            {
                // Кодировка для Unicode: Basic Latin и Cyrillic
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
            };
            // Cериализация
            return JsonSerializer.Serialize(this, options);
        }
    }
}
