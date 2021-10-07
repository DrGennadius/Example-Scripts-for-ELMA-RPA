using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    public class CorrectRowInfoItem
    {
        public RowInfoItem OriginRowInfoItem { get; set; }

        public RowInfoItem NewRowInfoItem { get; set; }

        public bool IsValid { get; set; } = false;

        public double SimilarCoef { get; set; } = 0.0;

        public bool IsAutoCorrected
        {
            get => OriginRowInfoItem != null
                && NewRowInfoItem != null
                && !NewRowInfoItem.Equals(OriginRowInfoItem);
        }
    }
}
