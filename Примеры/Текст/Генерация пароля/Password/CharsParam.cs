using System.Text.Json.Serialization;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Параметры групп символов для генерации пароля.
    /// </summary>
    public class CharsParam
    {
        /// <summary>
        /// Символы.
        /// </summary>
        public char[] Chars { get; set; }

        /// <summary>
        /// Нормализированная вероятность появления символов текущей группы <c>Chars</c> в пароле.
        /// </summary>
        [JsonIgnore]
        public double NormalizedProbability { get; set; }

        /// <summary>
        /// Вероятность (или правильнее степень) появления символов текущей группы <c>Chars</c> в пароле.
        /// Чем больше число, тем чаще будут появляться символы текущей группы относительно других групп.
        /// </summary>
        public double Probability { get; set; }
    }
}
