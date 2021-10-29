using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Конфигурация пароля.
    /// </summary>
    public class PasswordConfiguration
    {
        /// <summary>
        /// Стандартная длина пароля.
        /// </summary>
        public const int DefaultPasswordLength = 8;

        /// <summary>
        /// Стандартные символы нижнего регистра.
        /// </summary>
        public const string DefaultPasswordCharsLowcase = "abcdefgijkmnopqrstwxyz";

        /// <summary>
        /// Стандартные символы верхнего регистра.
        /// </summary>
        public const string DefaultPasswordCharsUppercase = "ABCDEFGHJKLMNPQRSTWXYZ";

        /// <summary>
        /// Стандартные символы чисел.
        /// </summary>
        public const string DefaultPasswordCharsNumeric = "123456789";

        /// <summary>
        /// Стандартные специальные символы.
        /// </summary>
        public const string DefaultPasswordCharsSpecial = "*$-+?_&=!%{}/";

        /// <summary>
        /// Длина пароля.
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// Использовать нижний регистр.
        /// </summary>
        public bool UseLowcase { get; set; }

        /// <summary>
        /// Использовать верхний регистр.
        /// </summary>
        public bool UseUppercase { get; set; }

        /// <summary>
        /// Использовать числа.
        /// </summary>
        public bool UseNumeric { get; set; }

        /// <summary>
        /// Использовать специальные символы.
        /// </summary>
        public bool UseSpecial { get; set; }

        public PasswordConfiguration()
        {
            Reset();
        }

        public PasswordConfiguration(string jsonString)
        {
            FromJsonString(jsonString);
        }

        public PasswordConfiguration(PasswordConfiguration passwordGeneratorConfiguration)
        {
            CopyFromOther(passwordGeneratorConfiguration);
        }

        /// <summary>
        /// Сбросить параметры.
        /// </summary>
        public void Reset()
        {
            Length = DefaultPasswordLength;
            UseLowcase = true;
            UseUppercase = true;
            UseNumeric = true;
            UseSpecial = true;
        }

        /// <summary>
        /// Получить нормализированные (накопительные) параметры групп символов.
        /// </summary>
        /// <returns></returns>
        public List<CharsParam> GetCumulativedCharsParams()
        {
            // Получить параметры групп символов и обязательно сортируем.
            var cumulativedCharsParams = GetCharsParams().OrderByDescending(x => x.NormalizedProbability).ToList();

            // Нормализуем вероятности.
            // Суть в том, что будем случайно генерировать число [0;1),
            // а потом брать по этому числу определенную группу, которая будет меньше.
            double cumulativeSum = 0.0;
            foreach (var item in cumulativedCharsParams)
            {
                cumulativeSum += item.NormalizedProbability;
                item.NormalizedProbability = cumulativeSum;
            }

            return cumulativedCharsParams;
        }

        /// <summary>
        /// Получить параметры групп символов.
        /// </summary>
        /// <returns></returns>
        public List<CharsParam> GetCharsParams()
        {
            List<CharsParam> charsParams = new();

            if (UseLowcase)
            {
                charsParams.Add(new CharsParam()
                {
                    Chars = DefaultPasswordCharsLowcase.ToCharArray(),
                    Probability = 3
                });
            }
            if (UseUppercase)
            {
                charsParams.Add(new CharsParam()
                {
                    Chars = DefaultPasswordCharsUppercase.ToCharArray(),
                    Probability = 3
                });
            }
            if (UseNumeric)
            {
                charsParams.Add(new CharsParam()
                {
                    Chars = DefaultPasswordCharsNumeric.ToCharArray(),
                    Probability = 2
                });
            }
            if (UseSpecial)
            {
                charsParams.Add(new CharsParam()
                {
                    Chars = DefaultPasswordCharsSpecial.ToCharArray(),
                    Probability = 1
                });
            }

            if (charsParams.Count > 0)
            {
                // Считаем вероятности.
                double summ = charsParams.Sum(x => x.Probability);
                charsParams.ForEach(p => p.NormalizedProbability = p.Probability / summ);
            }

            return charsParams;
        }

        /// <summary>
        /// Конвертировать в строку в формате Json.
        /// </summary>
        /// <returns></returns>
        public string ToJsonString()
        {
            return JsonSerializer.Serialize(this);
        }

        /// <summary>
        /// Получить параметры из строки формата Json.
        /// </summary>
        /// <param name="jsonString"></param>
        public void FromJsonString(string jsonString)
        {
            var configuration = JsonSerializer.Deserialize<PasswordConfiguration>(jsonString);
            CopyFromOther(configuration);
        }

        /// <summary>
        /// Копировать значения параметров из другого экземпляра параметров пароля.
        /// </summary>
        /// <param name="passwordGeneratorConfiguration"></param>
        public void CopyFromOther(PasswordConfiguration passwordGeneratorConfiguration)
        {
            Length = passwordGeneratorConfiguration.Length;
            UseLowcase = passwordGeneratorConfiguration.UseLowcase;
            UseUppercase = passwordGeneratorConfiguration.UseUppercase;
            UseNumeric = passwordGeneratorConfiguration.UseNumeric;
            UseSpecial = passwordGeneratorConfiguration.UseSpecial;
        }
    }
}
