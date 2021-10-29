using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    public class LevenshteinHelper
    {
        public static int DeletionCost { get; set; } = 1;
        public static int InsertionCost { get; set; } = 1;

        public static int GetMinimum(int a, int b, int c) => (a = a < b ? a : b) < c ? a : c;

        public static int GetDistance(string firstWord, string secondWord)
        {
            var n = firstWord.Length + 1;
            var m = secondWord.Length + 1;
            var matrixD = new int[n, m];

            for (var i = 0; i < n; i++)
            {
                matrixD[i, 0] = i;
            }

            for (var j = 0; j < m; j++)
            {
                matrixD[0, j] = j;
            }

            for (var i = 1; i < n; i++)
            {
                for (var j = 1; j < m; j++)
                {
                    var substitutionCost = firstWord[i - 1] == secondWord[j - 1] ? 0 : 1;

                    matrixD[i, j] = GetMinimum(matrixD[i - 1, j] + DeletionCost,          // удаление
                                               matrixD[i, j - 1] + InsertionCost,         // вставка
                                               matrixD[i - 1, j - 1] + substitutionCost); // замена
                }
            }

            return matrixD[n - 1, m - 1];
        }
    }
}
