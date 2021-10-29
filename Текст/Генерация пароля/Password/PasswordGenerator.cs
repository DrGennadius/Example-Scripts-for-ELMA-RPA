using System;
using System.Collections.Generic;
using System.Linq;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Генератор паролей.
    /// </summary>
    public class PasswordGenerator
    {
        private readonly Random Random = new();

        /// <summary>
        /// Конфигурация пароля.
        /// </summary>
        public PasswordConfiguration Configuration { get; set; }

        /// <summary>
        /// Сгенерировать пароль определенной длины.
        /// </summary>
        /// <param name="length"></param>
        /// <returns></returns>
        public string Generate(int length)
        {
            return Generate(length, length);
        }

        /// <summary>
        /// Сгенерировать пароль 
        /// </summary>
        /// <param name="minLength"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        public string Generate(int minLength, int maxLength)
        {
            if (minLength <= 0 || maxLength <= 0 || minLength > maxLength)
            {
                return null;
            }

            if (Configuration == null)
            {
                Configuration = new PasswordConfiguration();
            }

            var cumulativedCharsParams = Configuration.GetCumulativedCharsParams();

            char[] password = BaseGenerate(cumulativedCharsParams, minLength, maxLength);

            password = DuplicateFilter(cumulativedCharsParams, password);

            password = GroupCorrection(cumulativedCharsParams, password);

            return new string(password);
        }

        /// <summary>
        /// Сгенерировать символ.
        /// </summary>
        /// <param name="charsParam"></param>
        /// <returns></returns>
        private char GenerateChar(CharsParam charsParam)
        {
            int subIndex = Random.Next(0, charsParam.Chars.Length - 1);
            return charsParam.Chars[subIndex];
        }

        /// <summary>
        /// Сгенерировать символ.
        /// </summary>
        /// <param name="charsParams"></param>
        /// <returns></returns>
        private char GenerateChar(List<CharsParam> charsParams)
        {
            var probability = Random.NextDouble();
            var charsParam = charsParams.SkipWhile(x => x.NormalizedProbability < probability).First();
            return GenerateChar(charsParam);
        }

        /// <summary>
        /// Сгенерировать символ, исключая указанный символ.
        /// </summary>
        /// <param name="charsParam"></param>
        /// <param name="excludeChar"></param>
        /// <returns></returns>
        private char GenerateCharExcludeChar(CharsParam charsParam, char excludeChar)
        {
            string charsString = new(charsParam.Chars);
            char[] chars = charsString.Replace(excludeChar + "", "").ToArray();
            int subIndex = Random.Next(0, chars.Length - 1);
            return chars[subIndex];
        }

        /// <summary>
        /// Сгенерировать символ, исключая указанный символ.
        /// </summary>
        /// <param name="charsParams"></param>
        /// <param name="excludeChar"></param>
        /// <returns></returns>
        private char GenerateCharExcludeChar(List<CharsParam> charsParams, char excludeChar)
        {
            var probability = Random.NextDouble();
            var charsParam = charsParams.SkipWhile(x => x.NormalizedProbability < probability).First();
            return GenerateCharExcludeChar(charsParam, excludeChar);
        }

        /// <summary>
        /// Базовая генерация пароля со случайной длиной в указаном промежутке.
        /// </summary>
        /// <param name="charsParams"></param>
        /// <param name="minLength"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        private char[] BaseGenerate(List<CharsParam> charsParams, int minLength, int maxLength)
        {
            if (maxLength == int.MaxValue)
            {
                maxLength--;
            }
            int length = minLength < maxLength ? Random.Next(minLength, maxLength + 1) : minLength;

            char[] password = new char[length];

            for (int i = 0; i < length; i++)
            {
                password[i] = GenerateChar(charsParams);
            }

            return password;
        }

        /// <summary>
        /// Фильтрация дублирующих подряд символов.
        /// </summary>
        /// <param name="charsParams"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private char[] DuplicateFilter(List<CharsParam> charsParams, char[] password)
        {
            if (password.Length == 0)
            {
                return password;
            }
            char lastChar = password[0];
            for (int i = 1; i < password.Length; i++)
            {
                if (lastChar == password[i])
                {
                    password[i] = GenerateCharExcludeChar(charsParams, lastChar);
                }
                lastChar = password[i];
            }

            return password;
        }

        /// <summary>
        /// Корректировка отсутствия уникальности групп символов.
        /// </summary>
        /// <param name="charsParams"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private char[] GroupCorrection(List<CharsParam> charsParams, char[] password)
        {
            var freeIndexes = GetNotFirstUniqueIndexes(charsParams, password);
            if (freeIndexes.Count == 0)
            {
                return password;
            }
            foreach (var group in charsParams)
            {
                bool isValid = false;
                foreach (var ch in password)
                {
                    if (group.Chars.Contains(ch))
                    {
                        isValid = true;
                        break;
                    }
                }
                if (!isValid)
                {
                    int randomIndex = Random.Next(0, freeIndexes.Count - 1);
                    int freeIndex = freeIndexes[randomIndex];
                    int subIndex = Random.Next(0, group.Chars.Length - 1);
                    password[freeIndex] = group.Chars[subIndex];
                    freeIndexes.RemoveAt(randomIndex);
                    if (freeIndexes.Count == 0)
                    {
                        break;
                    }
                }
            }

            return password;
        }

        /// <summary>
        /// Получить не первые уникальные индексы.
        /// </summary>
        /// <param name="charsParams"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private List<int> GetNotFirstUniqueIndexes(List<CharsParam> charsParams, char[] password)
        {
            List<int> mirrorIndexes = new();
            List<int> indexes = new();
            if (password.Length == 0)
            {
                return mirrorIndexes;
            }

            foreach (var group in charsParams)
            {
                for (int i = 0; i < password.Length; i++)
                {
                    if (group.Chars.Contains(password[i]))
                    {
                        indexes.Add(i);
                        break;
                    }
                }
            }

            for (int i = 0; i < password.Length; i++)
            {
                if (!indexes.Contains(i))
                {
                    mirrorIndexes.Add(i);
                }
            }

            return mirrorIndexes;
        }
    }
}
