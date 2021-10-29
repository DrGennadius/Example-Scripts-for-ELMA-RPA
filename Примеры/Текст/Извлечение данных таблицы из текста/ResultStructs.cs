using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    internal struct CheckHeaderCellByPatternResult
    {
        public bool IsSucces;

        public string CellText;

        public int BeginRowIndex;

        public int BeginColumnIndex;

        public int EndRowIndex;

        public int EndColumnIndex;

        public CheckHeaderCellByPatternResult(bool isSucces, string cellText, int beginRowIndex, int beginColumnIndex, int endRowIndex, int endColumnIndex)
        {
            IsSucces = isSucces;
            CellText = cellText;
            BeginRowIndex = beginRowIndex;
            BeginColumnIndex = beginColumnIndex;
            EndRowIndex = endRowIndex;
            EndColumnIndex = endColumnIndex;
        }
    }

    internal struct CalcHeaderCellsByPatternsResult
    {
        public bool IsSucces;

        public CheckHeaderCellByPatternResult[] CheckHeaderCellByPatternResults;

        public CalcHeaderCellsByPatternsResult(bool isSucces, CheckHeaderCellByPatternResult[] checkHeaderCellByPatternResults)
        {
            IsSucces = isSucces;
            CheckHeaderCellByPatternResults = checkHeaderCellByPatternResults;
        }
    }
}
