using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;

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
            // Получить имя хоста очень просто:
            context.HostName = Dns.GetHostName();

            // Далее рассмотрим 2 варианта получения IP в локальной сети.

            // Вариант 1 (через dns).
            // Проверяем доступны ли нам вообще подключения.
            if (NetworkInterface.GetIsNetworkAvailable())
            {
                // Можно вместо context.HostName использовать сразу Dns.GetHostName()
                IPHostEntry host = Dns.GetHostEntry(context.HostName);

                // Получаем первый экземпляр класс IPAddress с версией IP 4.
                var ipAddress = host
                    .AddressList
                    .FirstOrDefault(ip => ip.AddressFamily == AddressFamily.InterNetwork);

                if (ipAddress != null)
                {
                    // Сохраняем IP.
                    context.HostIP = ipAddress.ToString();
                }
            }

            // Вариант 2 (используем сокет).
            // Это более точный способ, когда на локальном компьютере доступно несколько IP-адресов.
            using Socket socket = new(AddressFamily.InterNetwork, SocketType.Dgram, 0);
            socket.Connect("8.8.8.8", 65530);
            IPEndPoint endPoint = socket.LocalEndPoint as IPEndPoint;
            context.HostIP = endPoint.Address.ToString();
        }
    }
}
