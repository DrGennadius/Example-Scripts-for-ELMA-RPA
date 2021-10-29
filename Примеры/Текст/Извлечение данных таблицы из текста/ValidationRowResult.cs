using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELMA.RPA.Scripts
{
    public class ValidationRowResult
    {
        public bool IsValid { get; set; }

        public CorrectRowInfoItem CorrectRowInfo { get; set; }
    }
}
