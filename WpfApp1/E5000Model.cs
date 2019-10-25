using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGCCExcelOp
{
    public class E5000Model
    {
        public List<SubstationModel> SubstationList { get; set; }

        public List<PowerPlantModel> PowerPlantList { get; set; }
    }

    public class SubstationModel
    {
        /// <summary>
        /// 变电所名字
        /// </summary>
        public string Name { get; set; }

        public string All { get; set; }

        public string Up { get; set; }

        public string Down { get; set; }
    }

    public class PowerPlantModel
    {
        public  string Name { get; set; }

        public string DailyElc { get; set; }
    }
}
