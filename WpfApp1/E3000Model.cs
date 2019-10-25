using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGCCExcelOp
{
    public class E3000Model
    {
        public string ZongjiaName { get; set; }

        public List<PowerTimePairsDto> PowerTimePairs { get; set; }
    }

    public class PowerTimePairsDto
    {
        public string TimeStr { get; set; }

        public string PowerStr { get; set; }
    }
}
