using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Экземпляр данного класса будет создан при выполнении скрипта.
    /// <summary>
    public class ScriptActivity
    {
        // Найти LCID можно в документе от сюда:
        // https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/70feba9f-294e-491e-b6eb-56532684c37f?redirectedfrom=MSDN

        private const uint WM_INPUTLANGCHANGEREQUEST = 0x0050;
        private const int HWND_BROADCAST = 0xffff;
        private const string en_US = "00000409";
        private const string ru_RU = "00000419";
        private const uint KLF_ACTIVATE = 1;

        /// <summary>
        /// Данная функция является точкой входа.
        /// <summary>
        public void Execute(Context context)
        {
            string LCID = en_US;
            if (string.IsNullOrEmpty(context.LanguageTag))
            {
                // Если не указываем, то делаем как бы переключение туда-суда ru<->en
                IntPtr HKL = GetKeyboardLayout(0);
                ushort keyboardId = (ushort)((uint)HKL >> 16);
                string hexValue = keyboardId.ToString("X");
                if (hexValue == "409")
                {
                    LCID = ru_RU;
                }
                // Проверять на ru-RU (419) не нужно, т.к. в самом начале уже установили en-US
            }
            else if (context.LanguageTag == "ru-RU")
            {
                LCID = ru_RU;
            }
            // Можно обойтись только этим, если нужно всегда переключать по конкретному LCID
            PostMessage(HWND_BROADCAST, WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, LoadKeyboardLayout(LCID, KLF_ACTIVATE));
        }

        [DllImport("user32.dll")]
        private static extern bool PostMessage(int hhwnd, uint msg, IntPtr wparam, IntPtr lparam);

        [DllImport("user32.dll")]
        private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);

        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);
    }
}
