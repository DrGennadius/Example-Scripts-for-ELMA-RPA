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
            // Инициализируем генератор. При необходимости можно сконфигурировать.
            PasswordGenerator passwordGenerator = new();
            // Сгенерировать пароль из 10 символов.
            context.Password = passwordGenerator.Generate(10);
        }
    }
}
