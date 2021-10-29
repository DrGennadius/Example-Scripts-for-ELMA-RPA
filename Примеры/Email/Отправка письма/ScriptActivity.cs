using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using MimeKit;
using MailKit.Net.Smtp;

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
            if (context.Attachments == null)
            {
                context.Attachments = new List<FileInfo>();
            }

            var message = new MimeMessage();

            // Можно использовать имя и email адрес отправителя/получателя таким обазом:
            //message.From.Add(new MailboxAddress("Имя отправителя", "email адрес отправителя"));
            //message.To.Add(new MailboxAddress("Имя получателя", "email адрес получателя"));
            // В этом случае нужно иметь отдельно переменные для имя и email
            
            // Для примера используем парсинг, а текст должен иметь такой формат "Имя <email>"
            message.From.Add(MailboxAddress.Parse(context.Sender));
            message.To.Add(MailboxAddress.Parse(context.Reciver));
            message.Subject = context.Subject;

            var builder = new BodyBuilder();
            builder.TextBody = context.MessageText;

            foreach (var attachment in context.Attachments)
            {
                // В примере используется FileInfo, если будут строки, то к FullName не нужно обращаться
                builder.Attachments.Add(attachment.FullName);
            }

            message.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                // В данном примере предполагается, что все данные для подключения
                // и аудентификации есть в контексте робота. 
                client.Connect(context.Host, (int)context.Port.Value, context.UseSsl.Value);

                // Заметка: это нужно, если сервер SMTP требует это
                client.Authenticate(context.UserName, context.Password);

                client.Send(message);
                client.Disconnect(true);
            }
        }
    }
}
