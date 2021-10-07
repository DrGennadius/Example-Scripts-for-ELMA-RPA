using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Общиий помощник для таблиц.
    /// </summary>
    public class CommonTableHelper
    {
        /// <summary>
        /// Это пустая строка.
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static bool IsEmptyRow(string[] row)
        {
            bool isEmpty = true;

            if (row.Length > 0)
            {
                foreach (var item in row)
                {
                    if (!string.IsNullOrWhiteSpace(item))
                    {
                        isEmpty = false;
                        break;
                    }
                }
            }

            return isEmpty;
        }
    }
}
