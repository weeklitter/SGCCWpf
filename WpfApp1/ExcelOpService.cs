using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGCCExcelOp
{
    /// <summary>
    /// 处理EXCEL表格服务类
    /// </summary>
    public class ExcelOpService
    {
        /// <summary>
        /// 处理E5000表格 将表格数据转换为dto
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public E5000Model GetE5000Model(ISheet sheet, string guangfu)
        {
            var e5000Model = new E5000Model
            {
                SubstationList = new List<SubstationModel>(),
                PowerPlantList = new List<PowerPlantModel>()
            };

            #region 变电所查询
            var row43 = sheet.GetRow(42);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "魏塘变",
                All = ConvertToDouble(row43.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row43.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row43.GetCell(13)).ToString("F2")
            });

            var row57 = sheet.GetRow(56);
            var row59 = sheet.GetRow(58);
            var lizebian = new SubstationModel
            {
                Name = "里泽变",
                All = (ConvertToDouble(row57.GetCell(21)) + ConvertToDouble(row59.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row57.GetCell(5)) + ConvertToDouble(row59.GetCell(5))).ToString("F2"),
                Down = (ConvertToDouble(row57.GetCell(13)) + ConvertToDouble(row59.GetCell(13))).ToString("F2")
            };
            e5000Model.SubstationList.Add(lizebian);

            var row55 = sheet.GetRow(54);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "下甸庙变",
                All = ConvertToDouble(row55.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row55.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row55.GetCell(13)).ToString("F2")
            });

            var row74 = sheet.GetRow(73);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "牛桥变",
                All = ConvertToDouble(row74.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row74.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row74.GetCell(13)).ToString("F2")
            });

            var row68 = sheet.GetRow(67);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "钱桥变",
                All = ConvertToDouble(row68.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row68.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row68.GetCell(13)).ToString("F2")
            });

            var row80 = sheet.GetRow(79);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "姚庄变",
                All = ConvertToDouble(row80.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row80.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row80.GetCell(13)).ToString("F2")
            });

            var row86 = sheet.GetRow(85);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "陶庄变",
                All = ConvertToDouble(row86.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row86.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row86.GetCell(13)).ToString("F2")
            });

            var row94 = sheet.GetRow(93);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "惠民变",
                All = ConvertToDouble(row94.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row94.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row94.GetCell(13)).ToString("F2")
            });

            var row100 = sheet.GetRow(99);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "杨庙变",
                All = ConvertToDouble(row100.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row100.GetCell(5)) + ConvertToDouble(row100.GetCell(9))).ToString("F2"),
                Down = ConvertToDouble(row100.GetCell(13)).ToString("F2")
            });

            var row106 = sheet.GetRow(105);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "亭桥变",
                All = ConvertToDouble(row106.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row106.GetCell(5)) + ConvertToDouble(row106.GetCell(9))).ToString("F2"),
                Down = ConvertToDouble(row106.GetCell(13)).ToString("F2")
            });

            var row112 = sheet.GetRow(111);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "洪溪变",
                All = ConvertToDouble(row112.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row112.GetCell(5)) + ConvertToDouble(row112.GetCell(9))).ToString("F2"),
                Down = ConvertToDouble(row112.GetCell(13)).ToString("F2")
            });

            var row118 = sheet.GetRow(117);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "西塘变",
                All = ConvertToDouble(row118.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row118.GetCell(5)) + ConvertToDouble(row118.GetCell(9))).ToString("F2"),
                Down = ConvertToDouble(row118.GetCell(13)).ToString("F2")
            });

            var row124 = sheet.GetRow(123);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "云寺变",
                All = ConvertToDouble(row124.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row124.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row124.GetCell(13)).ToString("F2")
            });

            var row126 = sheet.GetRow(125);
            var row128 = sheet.GetRow(127);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "丁栅变",
                All = (ConvertToDouble(row126.GetCell(21)) + ConvertToDouble(row128.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row126.GetCell(5)) + ConvertToDouble(row128.GetCell(5))).ToString("F2"),
                Down = (ConvertToDouble(row126.GetCell(13)) + ConvertToDouble(row128.GetCell(13))).ToString("F2"),
            });

            var row138 = sheet.GetRow(137);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "范泾变",
                All = ConvertToDouble(row138.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row138.GetCell(5)) + ConvertToDouble(row138.GetCell(9))).ToString("F2"),
                Down = ConvertToDouble(row138.GetCell(13)).ToString("F2")
            });

            var row26 = sheet.GetRow(25);
            var row28 = sheet.GetRow(27);
            var row30 = sheet.GetRow(29);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "嘉善变35kV",
                All = (ConvertToDouble(row26.GetCell(21)) + ConvertToDouble(row28.GetCell(21)) + ConvertToDouble(row30.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row26.GetCell(5)) + ConvertToDouble(row28.GetCell(5)) + ConvertToDouble(row30.GetCell(5))).ToString("F2"),
                Down = (ConvertToDouble(row26.GetCell(13)) + ConvertToDouble(row28.GetCell(13)) + ConvertToDouble(row30.GetCell(13))).ToString("F2")
            });

            var row61 = sheet.GetRow(60);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "晋亿变",
                All = ConvertToDouble(row61.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row61.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row61.GetCell(13)).ToString("F2")
            });

            var row34 = sheet.GetRow(33);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "万泰变",
                All = ConvertToDouble(row34.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row34.GetCell(5)) + ConvertToDouble(row34.GetCell(17))).ToString("F2"),
                Down = ConvertToDouble(row34.GetCell(13)).ToString("F2")
            });

            var row24 = sheet.GetRow(23);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "东云变35kV",
                All = ConvertToDouble(row24.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row24.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row24.GetCell(13)).ToString("F2")
            });

            var row32 = sheet.GetRow(31);
            var row33 = sheet.GetRow(32);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "电气化铁路",
                All = (ConvertToDouble(row32.GetCell(21)) + ConvertToDouble(row33.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row32.GetCell(5)) + ConvertToDouble(row32.GetCell(9)) +
                      ConvertToDouble(row33.GetCell(5)) + ConvertToDouble(row33.GetCell(9))).ToString("F2"),
                Down = (ConvertToDouble(row32.GetCell(13)) + ConvertToDouble(row33.GetCell(13))).ToString("F2")
            });

            var row49 = sheet.GetRow(48);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "智果变",
                All = ConvertToDouble(row49.GetCell(21)).ToString("F2"),
                Up = (ConvertToDouble(row49.GetCell(5)) + ConvertToDouble(row49.GetCell(9))).ToString("F2"),
                Down = ConvertToDouble(row49.GetCell(13)).ToString("F2")
            });

            var row144 = sheet.GetRow(143);
            var row146 = sheet.GetRow(145);
            var row148 = sheet.GetRow(147);
            var row150 = sheet.GetRow(149);
            var xinlunbianUp = ConvertToDouble(row144.GetCell(5)) + ConvertToDouble(row146.GetCell(5)) + ConvertToDouble(row148.GetCell(5)) + ConvertToDouble(row150.GetCell(5))
                + ConvertToDouble(row144.GetCell(9)) + ConvertToDouble(row146.GetCell(9)) + ConvertToDouble(row148.GetCell(9)) + ConvertToDouble(row150.GetCell(9));
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "星轮变20kV",
                All = (ConvertToDouble(row144.GetCell(21)) + ConvertToDouble(row146.GetCell(21)) +
                      ConvertToDouble(row148.GetCell(21)) + ConvertToDouble(row150.GetCell(21))).ToString("F2"),
                Up = xinlunbianUp.ToString("F2"),
                Down = (ConvertToDouble(row144.GetCell(13)) + ConvertToDouble(row146.GetCell(13)) +
                      ConvertToDouble(row148.GetCell(13)) + ConvertToDouble(row150.GetCell(13))).ToString("F2"),
            });

            var row140 = sheet.GetRow(139);
            var row142 = sheet.GetRow(141);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "昱辉变",
                All = (ConvertToDouble(row140.GetCell(21)) + ConvertToDouble(row142.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row140.GetCell(5)) + ConvertToDouble(row142.GetCell(5))).ToString("F2"),
                Down = (ConvertToDouble(row140.GetCell(13)) + ConvertToDouble(row142.GetCell(13))).ToString("F2")
            });

            var row165 = sheet.GetRow(164);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "宝林变",
                All = ConvertToDouble(row165.GetCell(21)).ToString("F2"),
                Up = ConvertToDouble(row165.GetCell(5)).ToString("F2"),
                Down = ConvertToDouble(row165.GetCell(13)).ToString("F2")
            });

            var row14 = sheet.GetRow(13);
            var row116 = sheet.GetRow(115);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "康鼎变",
                All = (ConvertToDouble(row14.GetCell(21)) - ConvertToDouble(row116.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row14.GetCell(5)) - ConvertToDouble(row116.GetCell(5)) - ConvertToDouble(row116.GetCell(9))).ToString("F2"),
                Down = (ConvertToDouble(row14.GetCell(13)) - ConvertToDouble(row116.GetCell(13))).ToString("F2")
            });


            var row10 = sheet.GetRow(9);
            var row12 = sheet.GetRow(11);
            e5000Model.SubstationList.Add(new SubstationModel
            {
                Name = "富通变",
                All = (ConvertToDouble(row10.GetCell(21)) + ConvertToDouble(row12.GetCell(21))).ToString("F2"),
                Up = (ConvertToDouble(row10.GetCell(5)) + ConvertToDouble(row12.GetCell(5))).ToString("F2"),
                Down = (ConvertToDouble(row10.GetCell(13)) + ConvertToDouble(row12.GetCell(13))).ToString("F2")
            });

            #endregion

            #region 电厂信息赋值
            double.TryParse(guangfu, out double guangfuDian);

            var row36 = sheet.GetRow(35);
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "中成电厂",
                DailyElc = ConvertToDouble(row36.GetCell(21)).ToString("F2")
            });

            var row158 = sheet.GetRow(157);
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "协联电厂",
                DailyElc = ConvertToDouble(row158.GetCell(21)).ToString("F2")
            });

            var row168 = sheet.GetRow(167);
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "洪峰电厂",
                DailyElc = ConvertToDouble(row168.GetCell(21)).ToString("F2")
            });

            var row131 = sheet.GetRow(130);
            var rabishQuantity = ConvertToDouble(row131.GetCell(21));
            var dianchagnQuantity = e5000Model.PowerPlantList.Sum(it => Convert.ToDouble(it.DailyElc));
            var allDianchang = rabishQuantity + dianchagnQuantity;
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "电厂电量",
                DailyElc = allDianchang.ToString("F2")
            });

            var row155 = sheet.GetRow(154);
            var shanghaidianlaing = ConvertToDouble(row155.GetCell(21));
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "上海电量",
                DailyElc = shanghaidianlaing.ToString("F2")
            });

            var jiaxingdianlaing = e5000Model.SubstationList.Sum(s => Convert.ToDouble(s.All));
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "嘉兴电量",
                DailyElc = jiaxingdianlaing.ToString("F2")
            });

            var allXian = allDianchang + shanghaidianlaing + jiaxingdianlaing + guangfuDian;
            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "全县电量（含光）",
                DailyElc = allXian.ToString("F2")
            });

            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "垃圾电厂",
                DailyElc = rabishQuantity.ToString("F2")
            });

            e5000Model.PowerPlantList.Add(new PowerPlantModel
            {
                Name = "光伏电量",
                DailyElc = guangfuDian.ToString("F2")
            });

            #endregion

            return e5000Model;
        }

        public double ConvertToDouble(ICell cell)
        {
            if (cell.CellType == CellType.Formula)
            {
                try
                {
                    HSSFFormulaEvaluator hssEv = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                    cell = hssEv.EvaluateInCell(cell);
                }
                catch (Exception)
                {
                    XSSFFormulaEvaluator xssEv = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                    cell = xssEv.EvaluateInCell(cell);
                }
            }

            var cellStr = cell.ToString();
            if (string.IsNullOrEmpty(cellStr))
                return 0;
            else
                return Convert.ToDouble(cellStr);
        }

        /// <summary>
        /// 对模板赋值E5000
        /// </summary>
        /// <param name="e5000Model"></param>
        /// <param name="endSheet"></param>
        /// <returns></returns>
        public ISheet SetE5000Value(E5000Model e5000Model, ISheet endSheet)
        {
            int stationIndex = 20;
            foreach (var station in e5000Model.SubstationList)
            {
                var row = endSheet.GetRow(stationIndex);
                var allCell = row.GetCell(11);
                allCell.SetCellValue(station.All);
                var upCell = row.GetCell(12);
                upCell.SetCellValue(station.Up);
                var downCell = row.GetCell(13);
                downCell.SetCellValue(station.Down);
                stationIndex++;
            }

            int dianchangIndex = 20;
            foreach (var dianchang in e5000Model.PowerPlantList)
            {
                var row = endSheet.GetRow(dianchangIndex);
                var dailyDianliang = row.GetCell(2);
                dailyDianliang.SetCellValue(dianchang.DailyElc);
                dianchangIndex++;
            }

            return endSheet;
        }

        /// <summary>
        /// 处理E3000表格 将表格数据转换为dto
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public List<E3000Model> GetE3000Model(ISheet sheet)
        {
            var e3000List = new List<E3000Model>();
            var jiaxingZonggong = new E3000Model
            {
                ZongjiaName = "嘉兴总加",
                PowerTimePairs = new List<PowerTimePairsDto>()
            };

            var row4 = sheet.GetRow(3);
            var row6 = sheet.GetRow(5);
            var row7 = sheet.GetRow(6);
            for (int i = 0; i < 24; i++)
            {
                jiaxingZonggong.PowerTimePairs.Add(new PowerTimePairsDto
                {
                    TimeStr = i.ToString() + ":00",
                    PowerStr = (ConvertToDouble(row4.GetCell(i + 2)) + ConvertToDouble(row6.GetCell(i + 2)) - ConvertToDouble(row7.GetCell(i + 2))).ToString("F2")
                });
            }
            e3000List.Add(jiaxingZonggong);

            var diancahngZonggong = new E3000Model
            {
                ZongjiaName = "电厂总加",
                PowerTimePairs = new List<PowerTimePairsDto>()
            };

            for (int i = 0; i < 24; i++)
            {
                diancahngZonggong.PowerTimePairs.Add(new PowerTimePairsDto
                {
                    TimeStr = i.ToString() + ":00",
                    PowerStr = ConvertToDouble(row7.GetCell(i + 2)).ToString("F2")
                });
            }
            e3000List.Add(diancahngZonggong);

            var ganyuZonggong = new E3000Model
            {
                ZongjiaName = "干俞总加",
                PowerTimePairs = new List<PowerTimePairsDto>()
            };

            for (int i = 0; i < 24; i++)
            {
                ganyuZonggong.PowerTimePairs.Add(new PowerTimePairsDto
                {
                    TimeStr = i.ToString() + ":00",
                    PowerStr = ConvertToDouble(row6.GetCell(i + 2)).ToString("F2")
                });
            }
            e3000List.Add(ganyuZonggong);

            var quanjuZonggong = new E3000Model
            {
                ZongjiaName = "全局总加",
                PowerTimePairs = new List<PowerTimePairsDto>()
            };

            for (int i = 0; i < 24; i++)
            {
                quanjuZonggong.PowerTimePairs.Add(new PowerTimePairsDto
                {
                    TimeStr = i.ToString() + ":00",
                    PowerStr = ConvertToDouble(row4.GetCell(i + 2)).ToString("F2")
                });
            }
            e3000List.Add(quanjuZonggong);

            return e3000List;
        }

        /// <summary>
        /// 对模板赋值E3000
        /// </summary>
        /// <param name="e5000Model"></param>
        /// <param name="endSheet"></param>
        /// <returns></returns>
        public ISheet SetE3000Value(List<E3000Model> e3000Model, ISheet endSheet, string guangfu, string guangfuTime,
            string quanwangFuhe, string quanwangTime, string jiaxingFuhe, string jiaxingTime)
        {
            var row39 = endSheet.GetRow(38);
            var row40 = endSheet.GetRow(39);

            var guangfu111 = row39.GetCell(4);
            var guangfu222 = row40.GetCell(4);
            guangfu111.SetCellValue(guangfu);
            guangfu222.SetCellValue(guangfuTime);

            var row41 = endSheet.GetRow(40);
            var row42 = endSheet.GetRow(41);
            var quanwang = row41.GetCell(1);
            var quanwang222 = row42.GetCell(1);

            quanwang.SetCellValue(quanwangFuhe);
            quanwang222.SetCellValue(quanwangTime);

            var jiaxing = row41.GetCell(4);
            var jiaxing222 = row42.GetCell(4);

            jiaxing.SetCellValue(jiaxingFuhe);
            jiaxing222.SetCellValue(jiaxingTime);

            foreach (var stationQ in e3000Model)
            {
                var sartIndex = 0;
                if (stationQ.ZongjiaName == "嘉兴总加")
                {
                    sartIndex = 4;

                    var upFuhe = row39.GetCell(3);
                    var upTime = row40.GetCell(3);
                    var maxInfo = stationQ.PowerTimePairs.OrderBy(p => Convert.ToDouble(p.PowerStr)).Last();
                    upFuhe.SetCellValue(maxInfo.PowerStr);
                    upTime.SetCellValue(maxInfo.TimeStr);
                }
                else if (stationQ.ZongjiaName == "电厂总加")
                {
                    sartIndex = 8;

                    var upFuhe = row39.GetCell(5);
                    var upTime = row40.GetCell(5);
                    var maxInfo = stationQ.PowerTimePairs.OrderBy(p => Convert.ToDouble(p.PowerStr)).Last();
                    upFuhe.SetCellValue(maxInfo.PowerStr);
                    upTime.SetCellValue(maxInfo.TimeStr);
                }
                else if (stationQ.ZongjiaName == "干俞总加")
                {
                    sartIndex = 12;

                    var upFuhe = row39.GetCell(6);
                    var upTime = row40.GetCell(6);
                    var maxInfo = stationQ.PowerTimePairs.OrderBy(p => Convert.ToDouble(p.PowerStr)).Last();
                    upFuhe.SetCellValue(maxInfo.PowerStr);
                    upTime.SetCellValue(maxInfo.TimeStr);
                }
                else if (stationQ.ZongjiaName == "全局总加")
                {
                    sartIndex = 16;

                    var upFuhe = row39.GetCell(1);
                    var upTime = row40.GetCell(1);
                    var maxInfo = stationQ.PowerTimePairs.OrderBy(p => Convert.ToDouble(p.PowerStr)).Last();
                    upFuhe.SetCellValue(maxInfo.PowerStr);
                    upTime.SetCellValue(maxInfo.TimeStr);
                }

                var row = endSheet.GetRow(sartIndex - 1);
                var time = 2;
                foreach (var power in stationQ.PowerTimePairs)
                {
                    var cell = row.GetCell(time);
                    cell.SetCellValue(power.PowerStr);
                    if (power.TimeStr == "11:00")
                    {
                        sartIndex += 2;
                        time = 1;
                        row = endSheet.GetRow(sartIndex - 1);
                    }
                    time++;
                }
            }

            var timeCell = endSheet.GetRow(1).GetCell(1);
            var timeStr = DateTime.Now.ToString("D");
            timeCell.SetCellValue(timeStr);
            return endSheet;
        }
    }
}
