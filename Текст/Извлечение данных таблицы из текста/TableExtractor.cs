using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Извлекатель таблицы из текста.
    /// </summary>
    public class TableExtractor
    {
        readonly TableDetector _tableDetector = new();

        private string[,] _data = new string[0,0];

        public TableExtractor() {}

        public TableExtractor(TableDetectFeatures detectFeatures)
        {
            _tableDetector = new(detectFeatures);
        }

        public TableExtractor(TableDetector detectFeatures)
        {
            _tableDetector = detectFeatures;
        }

        public string[,] Data => _data;

        public bool Extract(string text)
        {
            var tableParameters = _tableDetector.Detect(text);
            return tableParameters.HasValue 
                && Extract(text, tableParameters.Value);
        }

        public bool Extract(string text, TableParameters tableParameters)
        {
            bool isSucces = true;

            int columnsLength = tableParameters.BeginColumnIndexes.Length;

            List<string[]> dataList = new();

            return isSucces;
        }
    }
}
