using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using MailKit.Net.Imap;
using MailKit;
using MailKit.Search;
using MimeKit;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Экземпляр данного класса будет создан при выполнении скрипта.
    /// <summary>
    public class ScriptActivity
    {
        /// <summary>
        /// Данная функция является точкой входа.
        /// <summary>
        public void Execute(Context context)
        {
            if (context.MessageContextStrings == null)
            {
                context.MessageContextStrings = new List<string>();
            }
            Directory.CreateDirectory(context.WorkDirectory);
            List<MessageContext> messageContexts = new List<MessageContext>();

            /*
             * Это лишь один из возможных вариантов. В текущем примере мы сохраняем
             * контекст писем как строку в формате JSON. Таким образом мы можем
             * связать несколько элементов в одном элементе.
             * Например, тело текста сообщения и если есть вложения, то и их тоже,
             * при этом это не разные списки и можно в другом месте соотнести
             * тело сообщения и вложения одного письма.
             */

            using (ImapClient client = new ImapClient())
            {
                // В данном примере предполагается, что все данные для подключения
                // и аудентификации есть в контексте робота. 
                client.Connect(context.Host, (int)context.Port.Value, context.UseSsl.Value);
                client.Authenticate(context.UserName, context.Password);

                // Бучем читать из стандартной папки "Входящие", "Inbox" и т.п.
                IMailFolder inbox = client.Inbox;
                // Нужны права и на запись для помечания прочитанных
                inbox.Open(FolderAccess.ReadWrite);

                // Важный для нас этап - настройка запроса (фильтрация)
                // 1. Нам нужны те, которые мы еще не прочитали и которые не удалены
                var query = SearchQuery.NotSeen.And(SearchQuery.NotDeleted);
                // 2. Далее нам нужно добавить поиск по тексту темы.
                if (!string.IsNullOrEmpty(context.SubjectText))
                {
                    query = query.And(SearchQuery.SubjectContains(context.SubjectText));
                }
                // 2. Далее нам нужно добавить поиск по тексту отправителя (например, email).
                if (!string.IsNullOrEmpty(context.SenderText))
                {
                    query = query.And(SearchQuery.HeaderContains("FROM", context.SenderText));
                }

                // Далее выполняем запрос и получаем уникальные идентификаторы писем,
                // которые должны соответветствовать запросу.
                IList<UniqueId> uids = client.Inbox.Search(query);

                // Проходимся по письмам и получаем текст с вложениями
                foreach (UniqueId uid in uids)
                {
                    string subDirectory = Path.Combine(context.WorkDirectory, uid.Id.ToString());
                    Directory.CreateDirectory(subDirectory);
                    MimeMessage message = client.Inbox.GetMessage(uid);
                    MessageContext messageContext = new MessageContext();
                    // Обращаем внимание, что тут можно указывать формат
                    messageContext.TextBody = message.GetTextBody(MimeKit.Text.TextFormat.Text);

                    foreach (MimeEntity attachment in message.Attachments)
                    {
                        var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                        fileName = Path.Combine(subDirectory, fileName);

                        using (var stream = File.Create(fileName))
                        {
                            if (attachment is MessagePart)
                            {
                                var mssagePart = (MessagePart)attachment;

                                mssagePart.Message.WriteTo(stream);
                            }
                            else
                            {
                                var part = (MimePart)attachment;

                                part.Content.DecodeTo(stream);
                            }
                        }

                        messageContext.AttachmentPaths.Add(fileName);
                    }

                    // Добавляем контекст в формате JSON строки.
                    context.MessageContextStrings.Add(messageContext.GetContextJSON());
                }

                // Помечаем как прочитанные. Не забываем,
                // что может понадобиться права на запись в папке
                inbox.AddFlags(uids, MessageFlags.Seen, true);
            }
        }
    }
}
