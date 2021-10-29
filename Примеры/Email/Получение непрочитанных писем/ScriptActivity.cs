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
            if (context.MessageTexts == null)
            {
                context.MessageTexts = new List<string>();
            }
            if (context.Attachments == null)
            {
                context.Attachments = new List<FileInfo>();
            }
            Directory.CreateDirectory(context.WorkDirectory);

            /*
             * Это лишь один из возможных вариантов. В текущем примере мы просто 
             * получаем списки текстов и вложений. Но они по сути никак не связаны.
             * Пока мы не можем использовать в контексте нестандартные типы в списке
             * и не можем хранить в одном элементе структуру, подсписки разного типа и т.д.
             * Но как вариант можно генерировать строку для одного письма, где будет
             * содержаться текст письма и пути к сохраненным файлам вложений и потом
             * в других местах парсить это.
             * Еще как вариант можно всю информацию сохранять в папке по каждому письму,
             * а в контекст хранить пути к папкам, далее работаем с этим.
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
                    // Обращаем внимание, что тут можно указывать формат
                    context.MessageTexts.Add(message.GetTextBody(MimeKit.Text.TextFormat.Text));

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

                        context.Attachments.Add(new FileInfo(fileName));
                    }
                }

                // Помечаем как прочитанные. Не забываем,
                // что может понадобиться права на запись в папке
                inbox.AddFlags(uids, MessageFlags.Seen, true);
            }
        }
    }
}
