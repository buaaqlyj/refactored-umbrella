using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;

using Statistics.Criterion.KV;
using Statistics.Criterion.Dose;
using Statistics.DataUtility;
using Statistics.Instrument.Standard;
using Statistics.Instrument.Tested;
using Statistics.Log;
using Statistics.IO;

namespace Statistics
{
    public static class JobMethods
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fullName"></param>
        /// <param name="strType"></param>
        /// <param name="strMacSerial"></param>
        /// <param name="strSensorSerial"></param>
        /// <param name="fis"></param>
        /// <param name="contin"></param>
        /// <returns></returns>
        public static FileInfo SearchForFile(string fullName, string strType, string strMacSerial, string strSensorSerial, FileInfo[] fis, out bool contin)
        {
            string name = "";
            string type = "", mac = "", sensor = "";
            string[] str = null;
            int rate = 0;
            bool hasSS1 = false, hasSS2 = true;

            ArrayList ignStrs = new ArrayList(new string[] { "30cc", "20cc" });
            MatchPattern tpM = MatchPattern.FullyMatch;
            MatchPattern msM = MatchPattern.FullyMatch;
            MatchPattern ssM = MatchPattern.FullyMatch;

            //全部完全匹配的情况
            name = strMacSerial + strSensorSerial;
            strMacSerial = strMacSerial.Replace("主机_", "").ToLower();
            strType = strType.ToLower();
            if (strSensorSerial.StartsWith("_探测器_"))
            {
                strSensorSerial = strSensorSerial.Replace("_探测器_", "").ToLower();
                hasSS1 = true;
            }
            foreach (FileInfo item in fis)
            {
                if (!item.Name.Contains(@"~$"))
                {
                    if (hasSS1 == (item.Name.Contains("_探测器_")))
                    {
                        if (item.Name.Contains(name) || name.Contains(item.Name))
                        {
                            //继续做合并
                            contin = true;
                            return item;
                        }
                    }
                }
            }


            string macPN = DataUtility.DataUtility.GetPureNumber(strMacSerial, false);
            string senPN = DataUtility.DataUtility.GetPureNumber(strSensorSerial, false);
            string _macPN = "";
            string _senPN = "";
            Dictionary<string, int> files = new Dictionary<string, int>();
            Dictionary<string, int> suspFiles = new Dictionary<string, int>();

            //详细排查其他的可能情况
            foreach (FileInfo item in fis)
            {
                if (!item.Name.Contains(@"~$"))
                {
                    str = Path.GetFileNameWithoutExtension(item.Name).Split(new string[] { "_主机_", "_探测器_" }, StringSplitOptions.RemoveEmptyEntries);
                    type = str[0];
                    if (type.Contains("_"))
                    {
                        type = type.Substring(type.LastIndexOf('_') + 1).ToLower();
                    }

                    mac = str[1].ToLower();
                    if (item.Name.Contains("_探测器_"))
                    {
                        sensor = str[2].ToLower();
                        hasSS2 = true;
                    }
                    else
                    {
                        sensor = "/";
                        hasSS2 = false;
                    }

                    _macPN = DataUtility.DataUtility.GetPureNumber(mac, false);
                    _senPN = DataUtility.DataUtility.GetPureNumber(sensor, false);

                    if (hasSS1 && hasSS2)
                    {
                        if (strSensorSerial == sensor)
                        {
                            if (ignStrs.Contains(strSensorSerial))
                            {
                                ssM = MatchPattern.DoNotNecessary;
                            }
                            else
                            {
                                ssM = MatchPattern.FullyMatch;
                            }
                        }
                        else if (_senPN == senPN)
                        {
                            if (ignStrs.Contains(strSensorSerial) && ignStrs.Contains(sensor))
                            {
                                ssM = MatchPattern.DoNotMatch;
                            }
                            else
                            {
                                ssM = MatchPattern.GeneralMatch;
                            }
                        }
                        else
                        {
                            ssM = MatchPattern.DoNotMatch;
                        }
                    }
                    else if (hasSS1 || hasSS2)
                    {
                        ssM = MatchPattern.DoNotMatch;
                    }
                    else
                    {
                        ssM = MatchPattern.DoNotNecessary;
                    }

                    if (strMacSerial == mac)
                    {
                        msM = MatchPattern.FullyMatch;
                    }
                    else if (_macPN == macPN)
                    {
                        msM = MatchPattern.GeneralMatch;
                    }
                    else
                    {
                        msM = MatchPattern.DoNotMatch;
                    }

                    rate = DataUtility.DataUtility.CompareStrings(type, strType);
                    //旧比较法：type.ToLower() == strType.ToLower()
                    if (TestedInstrument.IsEqualTo(type, strType))
                    {
                        tpM = MatchPattern.FullyMatch;
                    }
                    else if (rate > 50)
                    {
                        tpM = MatchPattern.GeneralMatch;
                    }
                    else
                    {
                        tpM = MatchPattern.DoNotMatch;
                    }

                    //分类处理
                    if (ssM > 0)
                    {
                        if (msM > 0)
                        {
                            files.Add(item.FullName, (int)(rate * (DataUtility.DataUtility.CompareStrings(strMacSerial, mac) + DataUtility.DataUtility.CompareStrings(_macPN, macPN)) * (DataUtility.DataUtility.CompareStrings(strSensorSerial, sensor) + DataUtility.DataUtility.CompareStrings(senPN, _senPN)) / 40000.0));
                        }
                        else if (tpM > 0)
                        {
                            suspFiles.Add(item.FullName, (int)(rate * (DataUtility.DataUtility.CompareStrings(strMacSerial, mac) + DataUtility.DataUtility.CompareStrings(_macPN, macPN)) * (DataUtility.DataUtility.CompareStrings(strSensorSerial, sensor) + DataUtility.DataUtility.CompareStrings(senPN, _senPN)) / 40000.0));
                        }
                    }
                    else if (((int)tpM * (int)msM) > 0)
                    {
                        if (ssM == MatchPattern.DoNotNecessary)
                        {
                            files.Add(item.FullName, (int)(rate * (DataUtility.DataUtility.CompareStrings(strMacSerial, mac) + DataUtility.DataUtility.CompareStrings(_macPN, macPN)) * (DataUtility.DataUtility.CompareStrings(strSensorSerial, sensor) + DataUtility.DataUtility.CompareStrings(senPN, _senPN)) / 40000.0));
                        }
                    }
                    else if (tpM == MatchPattern.DoNotMatch && msM > 0)
                    {
                        suspFiles.Add(item.FullName, (int)(rate * (DataUtility.DataUtility.CompareStrings(strMacSerial, mac) + DataUtility.DataUtility.CompareStrings(_macPN, macPN)) * (DataUtility.DataUtility.CompareStrings(strSensorSerial, sensor) + DataUtility.DataUtility.CompareStrings(senPN, _senPN)) / 40000.0));
                    }
                    ////仪器类型精确匹配，主机编号不匹配，探测器编号无
                    ////tpM = 2, msM = 0, ssM = 3
                    //if ((int)msM * (int)ssM > 0)
                    //{
                    //    //主机编号和探测器编号均匹配（模糊匹配或者完全匹配）
                    //    if (tpM > 0)
                    //    {
                    //        files.Add(item.FullName, rate);
                    //    } 
                    //    else
                    //    {
                    //        suspFiles.Add(item.FullName, (int)(rate * (CompareStrings(strMacSerial, mac) + CompareStrings(_macPN, macPN)) * (CompareStrings(strSensorSerial, sensor) + CompareStrings(senPN, _senPN)) / 40000.0));
                    //    }
                    //}
                    //else if (((int)msM + (int)ssM != 0) && tpM > 0 && (int)ssM != 3)
                    //{
                    //    //主机编号和探测器编号有一个不匹配，但仪器类型匹配
                    //    suspFiles.Add(item.FullName, (int)(rate * (CompareStrings(strMacSerial, mac) + CompareStrings(_macPN, macPN)) * (CompareStrings(strSensorSerial, sensor) + CompareStrings(senPN, _senPN)) / 40000.0));
                    //}
                }
            }

            rate = 0;
            if (files.Count > 0)
            {
                string filename = "";
                bool firsttime = true;

                foreach (KeyValuePair<string, int> item in files)
                {
                    if (item.Value > rate || firsttime)
                    {
                        filename = item.Key;
                        rate = item.Value;
                    }
                    firsttime = false;
                }
                //有模糊匹配项，返回相似概率最高的，继续做合并
                contin = true;
                return new FileInfo(filename);
            }
            else if (suspFiles.Count > 0)
            {
                ReportSuspiciousFiles(suspFiles, fullName);
                //没有匹配项，但有可疑项，不作处理
                contin = false;
                return null;
            }
            else
            {
                //没有可疑项和匹配项，复制新文档
                contin = true;
                return null;
            }
        }

        public static FileInfo[] GetFilesFromType(string pathBase, string type, string extension, out bool checkClear)
        {
            string[] paths = null;
            if (type.StartsWith("pmx"))
            {
                type = type.ToLower();
                paths = new string[] { "PMX-I", "PMX-III" };
            }
            else if (type.StartsWith("piranha"))
            {
                type = type.ToLower();
                paths = new string[] { "Piranha", "Piranha555", "Piranha657" };
            }
            else if (type.StartsWith("solidose"))
            {
                type = type.ToLower();
                paths = new string[] { "Solidose", "Solidose300", "Solidose308", "Solidose400" };
            }
            else if (type.StartsWith("35050a"))
            {
                type = type.ToLower();
                paths = new string[] { "35050A", "35050AT" };
            }
            else
            {
                paths = new string[] { type };
            }

            string path = "";
            FileInfo[] files = null;
            FileInfo[] allfiles = new FileInfo[0];
            foreach (string item in paths)
            {
                path = Path.Combine(pathBase, item);
                if (!Directory.Exists(path))
                {
                    Log.LogHelper.AddException("无法识别的仪器类型：" + item, true);
                    checkClear = false;
                    return null;
                }
                files = (new DirectoryInfo(path)).GetFiles(extension, SearchOption.AllDirectories);
                if (files.Length > 0)
                {
                    allfiles = DataUtility.DataUtility.CombineFileInfoArray(allfiles, files);
                }
            }
            checkClear = true;
            return allfiles;
        }
        /// <summary>
        /// 找模板页，加入
        /// 
        /// 搬运内容：
        /// 01、L2      02    12    证书编号
        /// 02、B4 -D4  04    02-04 送校单位
        /// 03、F4      04    06    联系地址
        /// 04、B5      05    02    仪器名称
        /// 05、F5      05    06    型号
        /// 06、H5      05    08    编号
        /// 07、J5      05    10    厂家
        /// 08、L5      05    12    电离室号
        /// 09、K7      07    11    温度
        /// 10、M7      07    13    湿度
        /// 11、J8      08    10    气压
        /// 12、D12-M12 12    04-13 量程
        /// 13、D18     18    04    数据
        /// 14、F18     18    06    数据
        /// 15、H18     18    08    数据
        /// 16、J18     18    10    数据
        /// 17、L18     18    13    数据
        /// 18、D19     19    04    数据
        /// 19、F19     19    06    数据
        /// 20、H19     19    08    数据
        /// 21、J19     19    10    数据
        /// 22、L19     19    12    数据
        /// 23、D20     20    04    数据
        /// 24、F20     20    06    数据
        /// 25、H20     20    08    数据
        /// 26、J20     20    10    数据
        /// 27、L20     20    12    数据
        /// 28、G31     31    07    记录者签名
        /// 29、I31     31    09    校对者签名
        /// 30、K31     31    11    日期
        /// </summary>
        /// <param name="sourceEx"></param>
        /// <param name="sourceIndex"></param>
        /// <param name="destiEx"></param>
        public static void CopyData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int pattern, string certIdori, bool needFix, bool shouldFix, int startDestiRowIndex, out int newSheetIndex, out bool success)
        {
            bool success1;
            bool noNeed = false;
            int templateIndex = -1;
            int startSourceRowIndex = -1;
            int destiIndex = -1;
            MSExcel.Range rr = null;
            MSExcel.Worksheet ws1 = null;

            bool checkClear;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();

            string temp;
            string text = "";
            string certId = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, "L2", out success1);
            if (!success1)
            {
                Log.LogHelper.AddException(@"无法提取到证书编号", true);
                success = false;
                newSheetIndex = -1;
                return;
            }

            foreach (MSExcel.Worksheet item in destiEx.ExcelWorkbook.Sheets)
            {
                if (item.Name == "统计")
                {
                    destiIndex = item.Index;
                }
                else if (item.Name == @"标准模板")
                {
                    //找到模板页
                    templateIndex = item.Index;
                }
                else if (item.Name.Contains(@"标准模板"))
                {
                    Log.LogHelper.AddDataError(@"第" + item.Index + "页发现多余的标准模板", true);
                }
                else
                {
                    temp = destiEx.GetText(destiEx.ExcelWorkbook, item.Index, "L2", out checkClear).Trim();
                    if (temp.StartsWith(@"20") && (temp.Length == 9 || temp.Length == 10))
                    {
                        //找到有证书编号的数据页
                        if (item.Name == certId)
                        {
                            if (FormOperator.MessageBox_Show_YesNo("在历史数据记录的Excel中已发现了证书编号为" + certId + "的页面，是否覆盖？选择是，进行覆盖。选择否，停止对本Excel的处理", "是否覆盖"))
                            {
                                ws1 = item;
                            }
                            else
                            {
                                Log.LogHelper.AddException(@"要合并入的数据已存在于第" + item.Index + "页", true);
                                newSheetIndex = -1;
                                success = false;
                                return;
                            }
                        }
                        else
                        {
                            exSheets.Add(item.Index, temp);
                        }
                    }
                }
            }

            if (!noNeed)
            {
                if (templateIndex == -1)
                {
                    Log.LogHelper.AddException(@"找不到数据的标准模板", true);
                    success = false;
                    newSheetIndex = -1;
                    return;
                }
                else
                {
                    //有模板页。没找到参考证书编号页
                    if (ws1 == null)
                    {
                        ws1 = (MSExcel.Worksheet)destiEx.ExcelWorkbook.Sheets[templateIndex];
                        if (exSheets.Count > 0)
                        {
                            //有有效数据页
                            if (destiIndex < 1)
                            {
                                //没有统计页，把有效数据页里最前面的序号赋给destiIndex
                                foreach (int item in exSheets.Keys)
                                {
                                    if (destiIndex < 1 || destiIndex > item)
                                    {
                                        destiIndex = item;
                                    }
                                }
                            }
                            foreach (int item in exSheets.Keys)
                            {
                                if (DataUtility.DataUtility.LargerThan(certId, exSheets[item]) && destiIndex < item)
                                {
                                    //在所有比参考编号小的页面里，挑一个最靠后的序号给destiIndex
                                    destiIndex = item;
                                }
                            }
                            //在destiIndex右侧复制模板页
                            ws1.Copy(Type.Missing, destiEx.ExcelWorkbook.Sheets[destiIndex]);
                            //把新复制的模板页赋给ws1
                            ws1 = (MSExcel.Worksheet)destiEx.ExcelWorkbook.Sheets[destiIndex + 1];
                        }
                        else
                        {
                            //没有有效数据页时，在模板页左侧复制模板页
                            ws1.Copy(destiEx.ExcelWorkbook.Sheets[templateIndex], Type.Missing);
                            //把原来模板页，现在的模板复制页的位置给了ws1
                            ws1 = (MSExcel.Worksheet)destiEx.ExcelWorkbook.Sheets[templateIndex];
                        }
                        if (!ws1.Name.Contains(@"标准模板"))
                        {
                            Log.LogHelper.AddException(@"标准模板复制出错", true);
                            success = false;
                            newSheetIndex = -1;
                            return;
                        }
                        if (!exSheets.ContainsValue(certId))
                        {
                            ws1.Name = certId;
                        }
                    }

                    newSheetIndex = ws1.Index;

                    //确定原始数据的数据行
                    for (int i = 15; i < 22; i++)
                    {
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i, 3, out checkClear).Trim();
                        if (text == "1")
                        {
                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i + 1, 3, out checkClear).Trim();
                            if (text == "2")
                            {
                                text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i + 2, 3, out checkClear).Trim();
                                if (text == "3")
                                {
                                    startSourceRowIndex = i;
                                    break;
                                }
                            }
                        }
                    }

                    if (startSourceRowIndex == -1)
                    {
                        Log.LogHelper.AddException(@"找不到原始数据所在的行", true);
                        success = false;
                        newSheetIndex = -1;
                        return;
                    }

                    //拷贝数据
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 2, 12, "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《证书编号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 4, 1, 4, 2, new string[] { @"送校单位：", @"单位名称：" }, "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《送校单位》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 4, 5, 4, 6, @"联系地址：", "@", out checkClear);
                    //if (!checkClear) AddException(@"《联系地址》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 1, 5, 2, @"仪器名称：", "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《仪器名称》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 5, 5, 6, @"型号：", "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《型号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 7, 5, 8, new string[] { @"主机编号：", @"编号：" }, "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《主机编号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 9, 5, 10, @"厂家：", "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《厂家》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 11, 5, 12, new string[] { @"探测器编号：", "电离室号：", "探测器号：" }, "@", out checkClear);
                    if (!checkClear) Log.LogHelper.AddException(@"《探测器编号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 31, 7, "", out checkClear);
                    //if (!checkClear) AddException(@"《记录者》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 31, 9, "", out checkClear);
                    //if (!checkClear) AddException(@"《校对者》数据复制错误", true);
                    CopyDate(sourceEx, sourceIndex, destiEx, ws1.Index, out checkClear);

                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 7, 11, "0.000", out checkClear);
                    if (needFix && !checkClear) Log.LogHelper.AddException(@"《温度》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 7, 13, "0.0%", out checkClear);
                    if (needFix && !checkClear) Log.LogHelper.AddException(@"《湿度》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 8, 10, "", out checkClear);
                    if (needFix && !checkClear) Log.LogHelper.AddException(@"《气压》数据复制错误", true);

                    rr = destiEx.GetRange(destiEx.ExcelWorkbook, ws1.Index, "M8", out checkClear);
                    if (needFix)
                    {
                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 13, "修正", out checkClear);
                    }
                    else
                    {
                        //电离室->半导体
                        rr = destiEx.GetRange(destiEx.ExcelWorkbook, ws1.Index, "L8", out checkClear);
                        rr.FormulaLocal = "";
                        rr.Formula = "";
                        rr.FormulaArray = "";
                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 12, "1.000000", "@", out checkClear);
                        if (shouldFix)
                        {
                            destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 13, "自修正", out checkClear);
                        }
                        else
                        {
                            destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 13, "不修正", out checkClear);
                        }
                    }

                    ArrayList dataStructList = new ArrayList();
                    DataStruct dataStruct;
                    int standardRowIndex = -1;
                    string keyword = "标准";
                    switch (pattern)
                    {
                        case 2:
                            keyword = "滤片";
                            break;
                    }
                    for (int p = 13; p < 18; p++)
                    {
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, p, 1, out checkClear).Trim() + sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, p, 2, out checkClear).Trim() + sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, p, 3, out checkClear).Trim();
                        if (text.ToLower().Contains(keyword))
                        {
                            standardRowIndex = p;
                            break;
                        }
                    }

                    switch (pattern)
                    {
                        case 0:
                            //Dose
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 27, 13, "@", out checkClear);
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, "", out checkClear);
                            if (!checkClear) Log.LogHelper.AddException(@"《量程》数据复制错误", true);

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out checkClear).Trim();
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 4, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 6, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 8, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 10, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 12, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 12, out checkClear).Trim();
                            if (text.StartsWith("单位"))
                            {
                                CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 13, "", out checkClear);
                                if (!checkClear) Log.LogHelper.AddException(@"《单位》数据复制错误", true);
                            }
                            else
                            {
                                dataStruct = DataStruct.CalDataRange(dataStructList, sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out checkClear).Trim(), pattern, out success);
                                switch (dataStruct.DataRanges)
                                {
                                    case DataRange.cGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "cGy", out checkClear);
                                        break;
                                    case DataRange.mGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mGy", out checkClear);
                                        break;
                                    case DataRange.mR:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mR", out checkClear);
                                        break;
                                    case DataRange.R:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "R", out checkClear);
                                        break;
                                    case DataRange.uGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "μGy", out checkClear);
                                        break;
                                    case DataRange.Unknown:
                                        Log.LogHelper.AddException("无法判断数据单位", true);
                                        break;
                                }
                                switch (dataStruct.Distance)
                                {
                                    case Distance.d1:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.0m", out checkClear);
                                        break;
                                    case Distance.d1_5:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.5m", out checkClear);
                                        break;
                                    case Distance.Unknown:
                                        Log.LogHelper.AddException("无法获取标准值，判断测试距离", true);
                                        break;
                                }
                            }

                            break;
                        case 1:
                            //CT
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 27, 13, "@", out checkClear);
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, "", out checkClear);
                            if (!checkClear) Log.LogHelper.AddException(@"《量程》数据复制错误", true);

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out checkClear).Trim();
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 4, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 6, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 8, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 10, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 12, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear));

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 12, out checkClear).Trim();
                            if (text.StartsWith("单位"))
                            {
                                CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 13, "", out checkClear);
                                if (!checkClear) Log.LogHelper.AddException(@"《单位》数据复制错误", true);
                            }
                            else
                            {
                                dataStruct = DataStruct.CalDataRange(dataStructList, sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out checkClear).Trim(), pattern, out success);
                                switch (dataStruct.DataRanges)
                                {
                                    case DataRange.mGycm:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mGycm", out checkClear);
                                        break;
                                    case DataRange.mGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mGy", out checkClear);
                                        break;
                                    case DataRange.cGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "cGy", out checkClear);
                                        break;
                                    case DataRange.mR:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mR", out checkClear);
                                        break;
                                    case DataRange.R:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "R", out checkClear);
                                        break;
                                    case DataRange.uGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "μGy", out checkClear);
                                        break;
                                    case DataRange.cGycm:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "cGycm", out checkClear);
                                        break;
                                    case DataRange.Rcm:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "Rcm", out checkClear);
                                        break;
                                    case DataRange.Unknown:
                                        Log.LogHelper.AddException("无法判断数据单位", true);
                                        break;
                                }
                                switch (dataStruct.Distance)
                                {
                                    case Distance.d1:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.0m", out checkClear);
                                        break;
                                    case Distance.d1_5:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.5m", out checkClear);
                                        break;
                                    case Distance.Unknown:
                                        Log.LogHelper.AddException("无法获取标准值，判断测试距离", true);
                                        break;
                                }
                            }
                            //CopyCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, startDestiRowIndex, "", out checkClear);
                            break;
                        case 2:
                            //KV
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 28, 13, "@", out checkClear);
                            //if (!checkClear) AddException(@"《备注》数据复制错误", true);
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, @"/", "", out checkClear);
                            if (!checkClear) Log.LogHelper.AddException(@"《量程》数据复制错误", true);

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out checkClear).Trim();
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 4, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 6, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 8, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 10, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 12, startDestiRowIndex, "", standardRowIndex, text, pattern, out checkClear);

                            //text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 12, out checkClear).Trim();
                            //if (text.StartsWith("单位"))
                            //{
                            //    //CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 13, "", out checkClear);
                            //    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, "", out checkClear);
                            //    if (!checkClear) AddException(@"《量程》数据复制错误", true);

                            //    if (!checkClear) AddException(@"《单位》数据复制错误", true);
                            //}
                            //else
                            //{
                            //    destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "KV", out checkClear);
                            //}

                            //text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 10, out checkClear).Trim();
                            //if (text.StartsWith("测试类型"))
                            //{
                            //    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 11, "", out checkClear);
                            //    if (!checkClear) AddException(@"《测试类型》数据复制错误", true);
                            //}
                            //else
                            //{
                            //    text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 5, 6, out checkClear).Trim().ToLower();
                            //    switch (text)
                            //    {
                            //        case "iba":
                            //            keyword = "ppV";
                            //            break;
                            //        case "radcal":
                            //            keyword = "kVp";
                            //            break;
                            //        default:
                            //            keyword = @"";
                            //            break;
                            //    }
                            //    destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 11, keyword, out checkClear);                                
                            //}
                            break;
                        default:
                            break;
                    }
                }
            }
            else
            {
                newSheetIndex = 1;
            }
            success = true;
        }

        public static void CopyData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int pattern, string certIdori, bool needFix, bool shouldFix, int startDestiRowIndex, out bool success)
        {
            int index = 0;
            CopyData(sourceEx, sourceIndex, destiEx, pattern, certIdori, needFix, shouldFix, startDestiRowIndex, out index, out success);
        }

        public static void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, string style, out bool sc)
        {
            bool checkClear;
            string text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, rowIndex, columnIndex, out checkClear).Trim();
            if (text == "")
            {
                sc = false;
                return;
            }
            if (style == "")
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, rowIndex, columnIndex, text, out checkClear);
            }
            else
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, rowIndex, columnIndex, text, style, out checkClear);
            }
            if (checkClear)
            {
                sc = true;
            }
            else
            {
                sc = false;
            }
        }

        public static void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, string style, bool checkDouble, out bool sc)
        {
            bool checkClear;
            double temp_double = 0.0;
            string text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, rowIndex, columnIndex, out checkClear).Trim();
            bool isDouble = double.TryParse(text, out temp_double);
            if (text == "" || !isDouble)
            {
                sc = false;
                return;
            }

            if (style == "")
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, rowIndex, columnIndex, text, out checkClear);
            }
            else
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, rowIndex, columnIndex, text, style, out checkClear);
            }
            if (checkClear)
            {
                sc = true;
            }
            else
            {
                sc = false;
            }
        }

        public static void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, string defaultValue, string style, out bool sc)
        {
            bool checkClear;
            string text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, rowIndex, columnIndex, out checkClear).Trim();
            if (text == "")
            {
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, rowIndex, columnIndex, defaultValue, out checkClear);
            }
            else if (style == "")
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, rowIndex, columnIndex, text, out checkClear);
            }
            else
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, rowIndex, columnIndex, text, style, out checkClear);
            }
            if (checkClear)
            {
                sc = true;
            }
            else
            {
                sc = false;
            }
        }

        public static void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, int new_row, int new_column, string pre, string style, out bool sc)
        {
            bool checkClear;
            string text = sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, sourceIndex, rowIndex, columnIndex, new_row, new_column, pre, out checkClear);
            if (text == "")
            {
                sc = false;
                return;
            }
            if (style == "")
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, new_row, new_column, text, out checkClear);
            }
            else
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, new_row, new_column, text, style, out checkClear);
            }
            if (checkClear)
            {
                sc = true;
            }
            else
            {
                sc = false;
            }
        }

        public static void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, int new_row, int new_column, string[] pre, string style, out bool sc)
        {
            bool checkClear;
            string text = sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, sourceIndex, rowIndex, columnIndex, new_row, new_column, pre, out checkClear);
            if (text == "")
            {
                sc = false;
                return;
            }
            if (style == "")
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, new_row, new_column, text, out checkClear);
            }
            else
            {
                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, new_row, new_column, text, style, out checkClear);
            }
            if (checkClear)
            {
                sc = true;
            }
            else
            {
                sc = false;
            }
        }

        public static void CopyThreeKVData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int startSourceRowIndex, int columnIndex, int startDestiRowIndex, string style, int standardRowIndex, string range, int pattern, out bool sc)
        {
            bool checkClear;
            double temp_double = 0.0;
            string text;
            MSExcel.Range rr = null;
            KVCriterion dataCri = KVCriterion.Null;

            sc = true;

            //规范
            rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(13, columnIndex), out checkClear);
            text = rr.Text.ToString().Trim();
            rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(13, columnIndex), out checkClear);
            rr.Value2 = text;

            //标准值
            sourceEx.GetCriterion(sourceEx.ExcelWorkbook, sourceIndex, columnIndex, true, out text, out dataCri);

            //仪器滤片
            if (standardRowIndex > 12)
            {
                rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(standardRowIndex, columnIndex), out checkClear);
                if (rr != null)
                {
                    text = rr.Text.ToString().Trim();
                    rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(15, columnIndex), out checkClear);
                    rr.Value2 = text;
                }
                else
                {
                    rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(15, columnIndex), out checkClear);
                    rr.Value2 = @"/";
                }
            }
            else
            {
                rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(15, columnIndex), out checkClear);
                rr.Value2 = @"/";
            }

            //复制数据
            if (style == "")
            {
                for (int i = 0; i < 3; i++)
                {
                    rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear);
                    if (rr.Value2 == null || string.IsNullOrWhiteSpace(rr.Value2.ToString()))
                    {
                        sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, @"/", out checkClear);
                        sourceEx.ExcelWorkbook.Save();
                        sourceEx.ExcelWorkbook.Saved = true;
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear).Trim();
                    }
                    else
                    {
                        text = rr.Value2.ToString().Trim();
                    }

                    if (double.TryParse(text, out temp_double) || text == @"/")
                    {
                        destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, columnIndex, text, out checkClear);
                    }
                    else
                    {
                        Log.LogHelper.AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                        sc = false;
                    }
                }
            }
            else
            {
                for (int i = 0; i < 3; i++)
                {
                    rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear);
                    if (rr.Value2 == null || string.IsNullOrWhiteSpace(rr.Value2.ToString()))
                    {
                        sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, @"/", out checkClear);
                        sourceEx.ExcelWorkbook.Save();
                        sourceEx.ExcelWorkbook.Saved = true;
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear).Trim();
                    }
                    else
                    {
                        text = rr.Value2.ToString().Trim();
                    }

                    if (double.TryParse(text, out temp_double) || text == @"/")
                    {
                        destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, columnIndex, text, style, out checkClear);
                    }
                    else
                    {
                        Log.LogHelper.AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                        sc = false;
                    }
                }
            }
        }

        public static DataStruct CopyThreeDoseData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int startSourceRowIndex, int columnIndex, int startDestiRowIndex, string style, int standardRowIndex, string range, int pattern, out bool sc)
        {
            bool checkClear;
            double temp_double = 0.0;
            double diValue = 0.00001;
            string text;
            MSExcel.Range rr = null;

            sc = true;

            //伦琴
            //text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, "D12", out checkClear).Trim().ToLower();
            //if (text.EndsWith(@" mr"))
            //{
            //    rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + 6, columnIndex, out checkClear);
            //    if ((bool)rr.HasFormula)
            //    {
            //        text = rr.FormulaLocal.ToString().Trim() + @"/33.97/0.000258";
            //        destiEx.WriteFormula(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + 6, columnIndex, text, out checkClear);
            //    }
            //}

            //规范
            rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(13, columnIndex), out checkClear);
            text = rr.Text.ToString().Trim();
            rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(13, columnIndex), DataUtility.DataUtility.PositionString(13, columnIndex + 1), out checkClear);
            rr.Value2 = text;

            //标准值
            if (standardRowIndex > 12)
            {
                rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(standardRowIndex, columnIndex), out checkClear);
                text = rr.Text.ToString().Trim();
                rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(16, columnIndex), DataUtility.DataUtility.PositionString(16, columnIndex + 1), out checkClear);
                rr.Value2 = text;
                if (double.TryParse(text, out temp_double))
                {
                    diValue = temp_double;
                }
            }

            //destiEx.WriteValue(destiIndex, 16, columnIndex, text, @"0.00000000_", out checkClear);

            //复制数据
            int count = 0;
            double daValue = 0;
            if (style == "")
            {
                for (int i = 0; i < 3; i++)
                {
                    rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear);
                    if (rr.Value2 == null || string.IsNullOrWhiteSpace(rr.Value2.ToString()))
                    {
                        sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, @"/", out checkClear);
                        sourceEx.ExcelWorkbook.Save();
                        sourceEx.ExcelWorkbook.Saved = true;
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear).Trim();
                    }
                    else
                    {
                        text = rr.Value2.ToString().Trim();
                    }

                    if (double.TryParse(text, out temp_double) || text == @"/")
                    {
                        destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, columnIndex, text, out checkClear);
                        if (text != @"/")
                        {
                            count++;
                            daValue += temp_double;
                        }
                    }
                    else
                    {
                        Log.LogHelper.AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                        sc = false;
                    }
                }

                if (count > 0)
                {
                    daValue /= count;
                    return new DataStruct(daValue, diValue, range, pattern);
                }
                else
                {
                    sc = false;
                    return null;
                }
            }
            else
            {
                for (int i = 0; i < 3; i++)
                {
                    rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear);
                    if (rr.Value2 == null || string.IsNullOrWhiteSpace(rr.Value2.ToString()))
                    {
                        sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, @"/", out checkClear);
                        sourceEx.ExcelWorkbook.Save();
                        sourceEx.ExcelWorkbook.Saved = true;
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear).Trim();
                    }
                    else
                    {
                        text = rr.Value2.ToString().Trim();
                    }

                    if (double.TryParse(text, out temp_double) || text == @"/")
                    {
                        destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, columnIndex, text, style, out checkClear);
                        if (text != @"/")
                        {
                            count++;
                            daValue += temp_double;
                        }
                    }
                    else
                    {
                        Log.LogHelper.AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                        sc = false;
                    }
                }

                if (count > 0)
                {
                    daValue /= count;
                    return new DataStruct(daValue, diValue, range, pattern);
                }
                else
                {
                    sc = false;
                    return null;
                }
            }
        }

        public static DataStruct CopyThreeCTData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int startSourceRowIndex, int columnIndex, int startDestiRowIndex, string style, int standardRowIndex, string range, int pattern, out bool sc)
        {
            bool checkClear;
            bool conti = true;
            double temp_double = 0.0;
            double diValue = 0.00001;
            string text;
            MSExcel.Range rr = null;
            NormalDoseCriterion dataCri = NormalDoseCriterion.Null;

            sc = true;

            conti = TestFlag(sourceEx, sourceIndex, startSourceRowIndex, columnIndex);
            if (conti)
            {
                //确定数据的规范
                if (sourceEx.GetCriterion(sourceEx.ExcelWorkbook, sourceIndex, columnIndex, true, out text, out dataCri))
                {
                    //规范
                    rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(13, columnIndex), out checkClear);
                    text = rr.Text.ToString().Trim();
                    //rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(13, sourceEx.GetColumnByCriterion(dataCri, 2)), DataUtility.DataUtility.PositionString(13, sourceEx.GetColumnByCriterion(dataCri, 2) + 1), out checkClear);
                    rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(13, columnIndex), DataUtility.DataUtility.PositionString(13, columnIndex + 1), out checkClear);
                    rr.Value2 = text;
                    //电压写入14行
                    destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, 14, columnIndex, dataCri.Voltage, out checkClear);
                    //半值层写入15行
                    destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, 15, columnIndex, dataCri.HalfValueLayer, out checkClear);

                    //标准值
                    if (standardRowIndex > 12)
                    {
                        rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(standardRowIndex, columnIndex), out checkClear);
                        text = rr.Text.ToString().Trim();
                        //rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(16, sourceEx.GetColumnByCriterion(dataCri, 2)), DataUtility.DataUtility.PositionString(16, sourceEx.GetColumnByCriterion(dataCri, 2) + 1), out checkClear);
                        rr = destiEx.GetRange(destiEx.ExcelWorkbook, destiIndex, DataUtility.DataUtility.PositionString(16, columnIndex), DataUtility.DataUtility.PositionString(16, columnIndex + 1), out checkClear);
                        rr.Value2 = text;
                        if (double.TryParse(text, out temp_double))
                        {
                            diValue = temp_double;
                        }
                        //destiEx.WriteValue(destiIndex, 16, columnIndex, text, @"0.00000000_", out checkClear);
                    }

                    //复制数据
                    int count = 0;
                    double daValue = 0;
                    if (style == "")
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear);
                            if (rr.Value2 == null || string.IsNullOrWhiteSpace(rr.Value2.ToString()))
                            {
                                sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, @"/", out checkClear);
                                sourceEx.ExcelWorkbook.Save();
                                sourceEx.ExcelWorkbook.Saved = true;
                                text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear).Trim();
                            }
                            else
                            {
                                text = rr.Value2.ToString().Trim();
                            }

                            if (double.TryParse(text, out temp_double) || text == @"/")
                            {
                                //destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, sourceEx.GetColumnByCriterion(dataCri, 2), text, out checkClear);
                                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, columnIndex, text, out checkClear);

                                if (text != @"/")
                                {
                                    count++;
                                    daValue += temp_double;
                                }
                            }
                            else
                            {
                                Log.LogHelper.AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                                sc = false;
                            }
                        }

                        if (count > 0)
                        {
                            daValue /= count;
                            return new DataStruct(daValue, diValue, range, pattern);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            rr = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear);
                            if (rr.Value2 == null || string.IsNullOrWhiteSpace(rr.Value2.ToString()))
                            {
                                sourceEx.WriteValue(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, @"/", out checkClear);
                                sourceEx.ExcelWorkbook.Save();
                                sourceEx.ExcelWorkbook.Saved = true;
                                text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, startSourceRowIndex + i, columnIndex, out checkClear).Trim();
                            }
                            else
                            {
                                text = rr.Value2.ToString().Trim();
                            }

                            if (double.TryParse(text, out temp_double) || text == @"/")
                            {
                                //destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, sourceEx.GetColumnByCriterion(dataCri, 2), text, "", out checkClear);
                                destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, startDestiRowIndex + i, columnIndex, text, style, out checkClear);

                                if (text != @"/")
                                {
                                    count++;
                                    daValue += temp_double;
                                }
                            }
                            else
                            {
                                Log.LogHelper.AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                                sc = false;
                            }
                        }

                        if (count > 0)
                        {
                            daValue /= count;
                            return new DataStruct(daValue, diValue, range, pattern);
                        }
                    }
                }
                else
                {
                    Log.LogHelper.AddException("规范内容有误：" + text, true);
                }
            }
            sc = false;
            return null;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceEx"></param>
        /// <param name="sourceIndex"></param>
        /// <param name="pattern"></param>
        /// <param name="shouldFix">需要修正但不需要进行修正操作</param>
        /// <param name="needFix">需要进行修正操作</param>
        /// <returns></returns>
        public static bool GetFixState(ExcelUtility sourceEx, int sourceIndex, int pattern, out bool needFix, out bool shouldFix)
        {
            if (pattern > 1)
            {
                needFix = false;
                shouldFix = false;
                return true;
            }
            bool checkClear;
            string tt1 = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, "M8", out checkClear);
            string tt2 = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, "L8", out checkClear);
            MSExcel.Range rr1 = sourceEx.GetRange(sourceEx.ExcelWorkbook, sourceIndex, "L8", out checkClear);
            double temp_double;
            bool isDouble = double.TryParse(tt2, out temp_double);
            if (tt1 == "修正")
            {
                needFix = true;
                shouldFix = true;
                return true;
            }
            else if (tt1.Contains("不修正") || tt2.Contains("不修正") || (tt1.Contains("不") && (tt1.Contains("修") || tt1.Contains("休"))) || (tt2.Contains("不") && (tt2.Contains("修") || tt2.Contains("休"))))
            {
                needFix = false;
                shouldFix = false;
                return true;
            }
            else if (tt1.Contains("自修正") || tt2.Contains("自修正"))
            {
                needFix = false;
                shouldFix = true;
                return true;
            }
            else if ((bool)rr1.HasFormula && isDouble)
            {
                needFix = true;
                shouldFix = true;
                return true;
            }
            else
            {
                Log.LogHelper.AddException("无法判断电离室与半导体类型", true);
                needFix = false;
                shouldFix = true;
                return false;
            }
        }
        /// <summary>
        /// 判断CT记录中3个数据位置的数据是否为有效的检定数据
        /// </summary>
        /// <param name="sourceEx"></param>
        /// <param name="sourceIndex"></param>
        /// <param name="startSourceRowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static bool TestFlag(ExcelUtility sourceEx, int sourceIndex, int startSourceRowIndex, int columnIndex)
        {
            Flag a, b, c;
            bool conti = false;
            double dig;
            bool checkClear;
            string text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(startSourceRowIndex, columnIndex), out checkClear).Trim();
            bool isDouble = double.TryParse(text, out dig);
            if (text == "/")
            {
                a = Flag.Zero;
            }
            else if (isDouble)
            {
                a = Flag.Pos;
            }
            else
            {
                a = Flag.Neg;
            }
            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(startSourceRowIndex + 1, columnIndex), out checkClear).Trim();
            isDouble = double.TryParse(text, out dig);
            if (text == "/")
            {
                b = Flag.Zero;
            }
            else if (isDouble)
            {
                b = Flag.Pos;
            }
            else
            {
                b = Flag.Neg;
            }
            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, DataUtility.DataUtility.PositionString(startSourceRowIndex + 2, columnIndex), out checkClear).Trim();
            isDouble = double.TryParse(text, out dig);
            if (text == "/")
            {
                c = Flag.Zero;
            }
            else if (isDouble)
            {
                c = Flag.Pos;
            }
            else
            {
                c = Flag.Neg;
            }

            if (((int)a + (int)b + (int)c > 0) && ((int)a * (int)b * (int)c > -1))
            {
                conti = true;
            }
            else
            {
                conti = false;
                if (((int)a + (int)b + (int)c + (int)a * (int)b * (int)c == 0) && !(a == Flag.Zero && b == Flag.Zero && c == Flag.Zero))
                {
                    Log.LogHelper.AddException("数据记录有未识别的数据", true);
                }
            }
            return conti;
        }

        public static void CopyDate(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, out bool sc)
        {
            bool checkClear;
            string text1, text2, text3, text4;
            string text = "";
            for (int i = 33; i > 26; i--)
            {
                text1 = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i, 10, out checkClear).Trim();
                text2 = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i, 11, out checkClear).Trim();
                text3 = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i, 12, out checkClear).Trim();
                text4 = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i, 13, out checkClear).Trim();
                text = text1 + text2 + text3 + text4;
                if (text != "")
                {
                    break;
                }
            }

            if (text == "")
            {
                sc = false;
                Log.LogHelper.AddException(@"《日期》数据复制错误：" + text, true);
                return;
            }
            else
            {
                text = text.Replace(@" ", "").Replace(@"年", " 年 ").Replace(@"月", " 月 ").Replace(@"日", " 日");
            }
            destiEx.WriteValue(destiEx.ExcelWorkbook, destiIndex, 31, 11, text, out checkClear);
            sc = true;
        }
        /// <summary>
        /// 统计，采集真实的规范列表
        /// </summary>
        /// <param name="eu"></param>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static bool Statistic(ExcelUtility eu, int pattern)
        {
            int stateIndex = -1;
            int lineIndex = 0;
            int limitPosition = 14;

            bool needInitialState = true;
            bool checkClear;
            bool firstTime = true;
            bool canPass = true, canPass1 = false;

            string insNumber = "";

            Dictionary<int, string> exSheets = null;
            Dictionary<int, string> cert = new Dictionary<int, string>();
            Dictionary<NormalDoseCriterion, int> existCriList = new Dictionary<NormalDoseCriterion, int>();

            //1.初始化：规范统计sheet结构，写入固定内容
            stateIndex = eu.InitialStatisticSheet(eu.ExcelWorkbook, out exSheets, out needInitialState, out existCriList, out limitPosition, out checkClear);
            if (!checkClear) { return false; }
            if (!needInitialState)
            {
                eu.WriteStateTitle(eu.ExcelWorkbook, pattern, stateIndex, existCriList, limitPosition);
                switch (pattern)
                {
                    case 0:
                        //1.搬运数据：先根据年份排序，然后复制数据
                        int[] indexDose = new int[exSheets.Count];
                        lineIndex = 0;
                        foreach (int item in exSheets.Keys)
                        {
                            indexDose[lineIndex] = item;
                            lineIndex++;
                        }
                        DataUtility.DataUtility.QuickSort(indexDose, exSheets, 0, indexDose.Length - 1);

                        lineIndex = 8;
                        for (int i = 0; i < exSheets.Count; i++)
                        {
                            cert.Add(lineIndex, exSheets[indexDose[i]]);
                            eu.CopyDoseOneYearData(eu.ExcelWorkbook, indexDose[i], stateIndex, lineIndex, exSheets[indexDose[i]], limitPosition, existCriList, out insNumber);

                            if (firstTime)
                            {
                                firstTime = false;
                            }
                            lineIndex++;
                        }
                        //2.分析数据

                        if (exSheets.Count > 1)
                        {
                            eu.WriteValue(eu.ExcelWorkbook, stateIndex, lineIndex, 1, @"重复性", out checkClear);
                            eu.ExcelWorkbook.Save();
                            eu.ExcelWorkbook.Saved = true;
                            //计算平均值和年稳定性
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR2（40kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 6, cert, "RQR3（50kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR4（60kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 10, cert, "RQR5（70kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR6（80kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 14, cert, "RQR7（90kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 16, cert, "RQR8（100kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 18, cert, "RQR9（120kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 20, cert, "RQR（140kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 22, cert, "RQR10（150kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                        }
                        break;
                    case 1:
                        //TODO: CT统计代码
                        //1.搬运数据：先根据年份排序，然后复制数据
                        //indexCT[数字序号]=页序号
                        //exSheets[页序号]=证书编号
                        int[] indexCT = new int[exSheets.Count];
                        lineIndex = 0;
                        foreach (int item in exSheets.Keys)
                        {
                            indexCT[lineIndex] = item;
                            lineIndex++;
                        }
                        DataUtility.DataUtility.QuickSort(indexCT, exSheets, 0, indexCT.Length - 1);

                        lineIndex = 8;
                        for (int i = 0; i < exSheets.Count; i++)
                        {
                            cert.Add(lineIndex, exSheets[indexCT[i]]);
                            //TODO:复制数据格式
                            eu.CopyCTOneYearData(eu.ExcelWorkbook, indexCT[i], stateIndex, i + 8, exSheets[indexCT[i]], limitPosition, existCriList, out insNumber);

                            if (firstTime)
                            {
                                firstTime = false;
                            }
                            lineIndex++;
                        }
                        //2.分析数据

                        if (exSheets.Count > 1)
                        {
                            eu.WriteValue(eu.ExcelWorkbook, stateIndex, lineIndex, 1, @"重复性", out checkClear);
                            eu.ExcelWorkbook.Save();
                            eu.ExcelWorkbook.Saved = true;
                            //计算平均值和年稳定性
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR2（40kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 6, cert, "RQR3（50kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR4（60kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 10, cert, "RQR5（70kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR6（80kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 14, cert, "RQR7（90kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 16, cert, "RQR8（100kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 18, cert, "RQR9（120kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 20, cert, "RQR（140kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                            canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 22, cert, "RQR10（150kV）", exSheets.Count > 1, true, -1);
                            canPass = canPass1 && canPass;
                        }
                        break;
                    case 2:
                        //TODO：KV统计代码
                        break;
                }
            }


            eu.ExcelWorkbook.Save();
            return canPass;
        }
        /// <summary>
        /// 统计，使用固定的规范列表
        /// </summary>
        /// <param name="eu"></param>
        /// <param name="pattern"></param>
        /// <param name="Perfect"></param>
        /// <param name="strCompany"></param>
        /// <param name="strType"></param>
        /// <param name="tempName"></param>
        /// <returns></returns>
        public static bool Statistic(ExcelUtility eu, int pattern, bool Perfect, string strCompany, string strType, string tempName, string certstr)
        {
            int stateIndex = -1;
            int lineIndex = 0;
            int limitPosition = 24;
            int newline = -1;

            bool needInitialState = true;
            bool checkClear;
            bool firstTime = true;
            bool canPass = true, canPass1 = false;

            string insNumber = "";

            Dictionary<int, string> exSheets = null;
            Dictionary<int, string> cert = new Dictionary<int, string>();
            Dictionary<NormalDoseCriterion, int> existDoseCriList = new Dictionary<NormalDoseCriterion, int>() { 
                { NormalDoseCriterion.RQR2_40  , NormalDoseCriterion.RQR2_40.Column  }, 
                { NormalDoseCriterion.RQR3_50  , NormalDoseCriterion.RQR3_50.Column  },
                { NormalDoseCriterion.RQR4_60  , NormalDoseCriterion.RQR4_60.Column  },
                { NormalDoseCriterion.RQR5_70  , NormalDoseCriterion.RQR5_70.Column  },
                { NormalDoseCriterion.RQR6_80  , NormalDoseCriterion.RQR6_80.Column  },
                { NormalDoseCriterion.RQR7_90  , NormalDoseCriterion.RQR7_90.Column  },
                { NormalDoseCriterion.RQR8_100 , NormalDoseCriterion.RQR8_100.Column },
                { NormalDoseCriterion.RQR9_120 , NormalDoseCriterion.RQR9_120.Column },
                { NormalDoseCriterion.RQR_140  , NormalDoseCriterion.RQR_140.Column  },
                { NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column}
            };

            Dictionary<KVCriterion, int> existKVCriList = new Dictionary<KVCriterion, int>() {
                { KVCriterion.RQR2_40  , KVCriterion.RQR2_40.Column  },
                { KVCriterion.RQR3_50  , KVCriterion.RQR3_50.Column  },
                { KVCriterion.RQR4_60  , KVCriterion.RQR4_60.Column  },
                { KVCriterion.RQR5_70  , KVCriterion.RQR5_70.Column  },
                { KVCriterion.RQR6_80  , KVCriterion.RQR6_80.Column  },
                { KVCriterion.RQR7_90  , KVCriterion.RQR7_90.Column  },
                { KVCriterion.RQR8_100 , KVCriterion.RQR8_100.Column },
                { KVCriterion.RQR9_120 , KVCriterion.RQR9_120.Column },
                { KVCriterion.RQR_140  , KVCriterion.RQR_140.Column  },
                { KVCriterion.RQR10_150, KVCriterion.RQR10_150.Column}
            };

            //1.初始化：规范统计sheet结构，写入固定内容
            stateIndex = eu.InitialStatisticSheet(eu.ExcelWorkbook, out exSheets, out needInitialState, out checkClear);
            if (!checkClear) { return false; }

            switch (pattern)
            {
                case 0:
                    if (needInitialState)
                    {
                        eu.WriteStateTitle(eu.ExcelWorkbook, pattern, stateIndex, existDoseCriList, limitPosition);
                        //写入送校单位
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 2, 2, strCompany, out checkClear);
                        //写入仪器名称
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 3, 2, strType, out checkClear);
                    }
                    //1.搬运数据：先根据年份排序，然后复制数据
                    int[] indexDose = new int[exSheets.Count];
                    foreach (int item in exSheets.Keys)
                    {
                        indexDose[lineIndex] = item;
                        lineIndex++;
                    }
                    DataUtility.DataUtility.QuickSort(indexDose, exSheets, 0, indexDose.Length - 1);

                    lineIndex = 8;
                    for (int i = 0; i < exSheets.Count; i++)
                    {
                        cert.Add(lineIndex, exSheets[indexDose[i]]);
                        eu.CopyDoseOneYearData(eu.ExcelWorkbook, indexDose[i], stateIndex, lineIndex, exSheets[indexDose[i]], limitPosition, existDoseCriList, out insNumber);

                        if (exSheets[indexDose[i]] == certstr)
                        {
                            newline = lineIndex;
                        }

                        if (firstTime && Perfect)
                        {
                            eu.DesiredName = tempName;

                            firstTime = false;
                        }
                        lineIndex++;
                    }

                    //2.分析数据
                    if (exSheets.Count > 1)
                    {
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, lineIndex, 1, @"重复性", out checkClear);
                        eu.ExcelWorkbook.Save();
                        eu.ExcelWorkbook.Saved = true;
                        //计算平均值和年稳定性
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR2（40kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 6, cert, "RQR3（50kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR4（60kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 10, cert, "RQR5（70kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR6（80kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 14, cert, "RQR7（90kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 16, cert, "RQR8（100kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 18, cert, "RQR9（120kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 20, cert, "RQR（140kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 22, cert, "RQR10（150kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                    }
                    break;
                case 1:
                    //TODO: CT统计代码
                    if (needInitialState)
                    {
                        eu.WriteStateTitle(eu.ExcelWorkbook, pattern, stateIndex, existDoseCriList, limitPosition);
                        //写入送校单位
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 2, 2, strCompany, out checkClear);
                        //写入仪器名称
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 3, 2, strType, out checkClear);
                    }
                    //1.搬运数据：先根据年份排序，然后复制数据
                    int[] indexCT = new int[exSheets.Count];
                    foreach (int item in exSheets.Keys)
                    {
                        indexCT[lineIndex] = item;
                        lineIndex++;
                    }
                    DataUtility.DataUtility.QuickSort(indexCT, exSheets, 0, indexCT.Length - 1);

                    lineIndex = 8;
                    for (int i = 0; i < exSheets.Count; i++)
                    {
                        cert.Add(lineIndex, exSheets[indexCT[i]]);
                        //TODO:复制数据格式
                        eu.CopyCTOneYearData(eu.ExcelWorkbook, indexCT[i], stateIndex, lineIndex, exSheets[indexCT[i]], limitPosition, existDoseCriList, out insNumber);

                        if (exSheets[indexCT[i]] == certstr)
                        {
                            newline = lineIndex;
                        }

                        if (firstTime && Perfect)
                        {
                            eu.DesiredName = tempName;

                            firstTime = false;
                        }
                        lineIndex++;
                    }

                    //2.分析数据
                    if (exSheets.Count > 1)
                    {
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, lineIndex, 1, @"重复性", out checkClear);
                        eu.ExcelWorkbook.Save();
                        eu.ExcelWorkbook.Saved = true;
                        //计算平均值和年稳定性
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR2（40kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 6, cert, "RQR3（50kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR4（60kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 10, cert, "RQR5（70kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR6（80kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 14, cert, "RQR7（90kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 16, cert, "RQR8（100kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 18, cert, "RQR9（120kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 20, cert, "RQR（140kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                        canPass1 = eu.StatisticsCTDoseOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 22, cert, "RQR10（150kV）", exSheets.Count > 1, true, newline);
                        canPass = canPass1 && canPass;
                    }
                    break;
                case 2:
                    //TODO：KV统计代码
                    limitPosition = 44;
                    if (needInitialState)
                    {
                        eu.WriteStateTitle(eu.ExcelWorkbook, pattern, stateIndex, existKVCriList, limitPosition);
                        //写入送校单位
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 2, 2, strCompany, out checkClear);
                        //写入仪器名称
                        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 3, 2, strType, out checkClear);
                    }

                    //1.搬运数据：先根据年份排序，然后复制数据
                    int[] indexKV = new int[exSheets.Count];
                    foreach (int item in exSheets.Keys)
                    {
                        indexKV[lineIndex] = item;
                        lineIndex++;
                    }
                    DataUtility.DataUtility.QuickSort(indexKV, exSheets, 0, indexKV.Length - 1);

                    lineIndex = 8;
                    for (int i = 0; i < exSheets.Count; i++)
                    {
                        cert.Add(lineIndex, exSheets[indexKV[i]]);
                        //TODO:复制数据格式
                        if (exSheets[indexKV[i]] == certstr)
                        {
                            eu.CopyKVOneYearData(eu.ExcelWorkbook, indexKV[i], stateIndex, lineIndex, exSheets[indexKV[i]], limitPosition, existKVCriList, true, out insNumber);
                        }
                        else
                        {
                            eu.CopyKVOneYearData(eu.ExcelWorkbook, indexKV[i], stateIndex, lineIndex, exSheets[indexKV[i]], limitPosition, existKVCriList, false, out insNumber);
                        }

                        if (exSheets[indexKV[i]] == certstr)
                        {
                            newline = lineIndex;
                        }

                        if (firstTime && Perfect)
                        {
                            eu.DesiredName = tempName;

                            firstTime = false;
                        }
                        lineIndex++;
                    }

                    //2.分析数据
                    eu.ExcelWorkbook.Save();
                    eu.ExcelWorkbook.Saved = true;
                    //计算平均值和年稳定性
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR2（40kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR3（50kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR4（60kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 16, cert, "RQR5（70kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 20, cert, "RQR6（80kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 24, cert, "RQR7（90kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 28, cert, "RQR8（100kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 32, cert, "RQR9（120kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 36, cert, "RQR（140kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    canPass1 = eu.StatisticsKVOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 40, cert, "RQR10（150kV）", exSheets.Count > 1, true, newline);
                    canPass = canPass1 && canPass;
                    break;
            }
            eu.ExcelWorkbook.Save();
            return canPass;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceEx"></param>
        /// <param name="stateIndex"></param>
        public static void TypeStandardize(ExcelUtility sourceEx, int stateIndex)
        {
            bool checkClear;
            string str = sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, stateIndex, 5, 5, 5, 6, @"型号：", out checkClear).Trim().ToLower();
            bool DoNotHave = true;
            if (str.Contains(@"unfors") || sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, stateIndex, 5, 9, 5, 10, @"厂家：", out checkClear).Trim().ToLower().Contains(@"unfors"))
            {
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, stateIndex, 5, 5, @"型号：", out checkClear);
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, stateIndex, 5, 6, @"Xi", out checkClear);
                sourceEx.ExcelWorkbook.Save();
                sourceEx.ExcelWorkbook.Saved = true;
                str = sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, stateIndex, 5, 5, 5, 6, @"型号：", out checkClear).Trim().ToLower();
            }
            if (str.Contains(@"baracuda"))
            {
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, stateIndex, 5, 5, @"型号：", out checkClear);
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, stateIndex, 5, 6, @"Barracuda", out checkClear);
                sourceEx.ExcelWorkbook.Save();
                sourceEx.ExcelWorkbook.Saved = true;
                str = sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, stateIndex, 5, 5, 5, 6, @"型号：", out checkClear).Trim().ToLower();
            }
            if (str.Contains(@"prianha"))
            {
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, stateIndex, 5, 5, @"型号：", out checkClear);
                sourceEx.WriteValue(sourceEx.ExcelWorkbook, stateIndex, 5, 6, @"Piranha", out checkClear);
                sourceEx.ExcelWorkbook.Save();
                sourceEx.ExcelWorkbook.Saved = true;
                str = sourceEx.GetMergeContent(sourceEx.ExcelWorkbook, stateIndex, 5, 5, 5, 6, @"型号：", out checkClear).Trim().ToLower();
            }
            foreach (string item in TestedInstrument.AllTypes)
            {
                if (item.ToLower() == str)
                {
                    DoNotHave = false;
                    break;
                }
            }
            if (DoNotHave)
            {
                Log.LogHelper.AddException(@"仪器类型可能出现手误", true);
                Log.LogHelper.AddLog("错误47", "  仪器类型：" + str, true);
            }
        }
        /// <summary>
        /// 生成证书和pdf记录
        /// </summary>
        /// <param name="excel">excel记录文件</param>
        /// <param name="sourceIndex">excel记录页索引号</param>
        /// <param name="wordPath">证书模板文件</param>
        /// <param name="savePath">证书记录存储文件夹</param>
        /// <param name="pdfPath">pdf记录存储文件夹</param>
        /// <param name="success">成功信号</param>
        public static void GenerateCert(ExcelUtility excel, int sourceIndex, int pattern, string wordPath, string savePath, string pdfPath, string tempFolder, bool shouldFix, out bool success)
        {
            //GenerateCert(_sr, stateIndex, path, pS.CertFolder, pS.PDFDataFolder, out success);
            try
            {
                WordUtility wu = new WordUtility(wordPath, out success);
                if (!success)
                {
                    Log.LogHelper.AddException("Word文档打开失败", true);
                    return;
                }
                string stemp1 = excel.GetText(excel.ExcelWorkbook, sourceIndex, "L2", out success);
                object otemp1;
                string wdName = "DYjl" + stemp1 + Path.GetExtension(wordPath);
                string pdfName = "DYjl" + stemp1 + "_" + excel.GetText(excel.ExcelWorkbook, sourceIndex, "B4", out success) + ".pdf";

                switch (pattern)
                {
                    case 0:
                        //Dose
                        //识别半导体和电离室，对证书模板做对应改动
                        if (shouldFix)
                        {
                            wu.WriteValue(wu.WordDocument, "M_JZJGR", "，并修正到标准温度、气压");
                        }
                        //1.普通复制（单位），剂量和CT独有
                        wu.WriteValue(wu.WordDocument, "M_DW", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M12", out success));
                        //2.读取小数点后三位（剂量值和校准因子），剂量和CT独有
                        wu.WriteDataValue(wu.WordDocument, "M_NC1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_LY1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L15", out success), "{0:F1}"); //小数点后一位
                        //3.普通复制（备注说明），剂量CT在B27，KV在B29
                        if (excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B27", out success) == null || excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B27", out success).ToString() == @"/")
                        {
                            wu.WriteValue(wu.WordDocument, "M_BZSM", "无");
                        }
                        else
                        {
                            wu.WriteValue(wu.WordDocument, "M_BZSM", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B27", out success));
                        }
                        break;
                    case 1:
                        //CT
                        //识别半导体和电离室，对证书模板做对应改动
                        if (shouldFix)
                        {
                            wu.WriteValue(wu.WordDocument, "M_JZTJ3", Environment.NewLine + "3、电离室戴保护管在辐射野中全照射。");
                            wu.WriteValue(wu.WordDocument, "M_JZJGR", "，并修正到标准温度、气压");
                        }
                        //备注说明普通复制，剂量CT在B27，KV在B29
                        if (excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B27", out success) == null || excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B27", out success).ToString() == @"/")
                        {
                            wu.WriteValue(wu.WordDocument, "M_BZSM", "无");
                        }
                        else
                        {
                            wu.WriteValue(wu.WordDocument, "M_BZSM", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B27", out success));
                        }
                        //校准因子单位写入
                        if (excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M27", out success) == null || excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M27", out success).ToString().ToLower() == "false")
                        {
                            wu.WriteValue(wu.WordDocument, "M_JZYZ", "无量纲");
                        }
                        else
                        {
                            wu.WriteValue(wu.WordDocument, "M_JZYZ", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M27", out success));
                        }
                        //1.普通复制（单位），剂量和CT独有
                        wu.WriteValue(wu.WordDocument, "M_DW", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M12", out success));
                        //2.读取小数点后三位（剂量值和校准因子），剂量和CT独有
                        wu.WriteDataValue(wu.WordDocument, "M_NC1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_NC5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L24", out success), "{0:F3}"); //小数点后三位
                        wu.WriteDataValue(wu.WordDocument, "M_LY1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J15", out success), "{0:F1}"); //小数点后一位
                        wu.WriteDataValue(wu.WordDocument, "M_LY5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L15", out success), "{0:F1}"); //小数点后一位
                        break;
                    case 2:
                        //TODO: KV
                        //1.普通复制（测试类型），KV独有
                        wu.WriteValue(wu.WordDocument, "M_CSLX", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "K12", out success));
                        //2.读取小数点后两位（实际峰值电压PPV和被测仪器显示值）
                        wu.WriteDataValue(wu.WordDocument, "M_PPV1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D14", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_PPV2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F14", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_PPV3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H14", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_PPV4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J14", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_PPV5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L14", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_VALUE1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D20", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_VALUE2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F20", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_VALUE3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H20", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_VALUE4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J20", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_VALUE5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L20", out success), "{0:F2}"); //小数点后两位
                        //3.两位小数的百分数（相对固有误差和过滤影响）
                        wu.WriteDataValue(wu.WordDocument, "M_XDGYWC1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D22", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_XDGYWC2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F22", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_XDGYWC3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H22", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_XDGYWC4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J22", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_XDGYWC5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L22", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX21", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX31", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "E82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX22", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX32", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "G82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX23", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX33", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "I82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX24", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX34", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "K82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX25", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L82", out success), "{0:0.00%}"); //两位小数的百分数
                        wu.WriteDataValue(wu.WordDocument, "M_GLYX35", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M82", out success), "{0:0.00%}"); //两位小数的百分数
                        //4.读取小数点后两位（辐照工作下限）
                        wu.WriteDataValue(wu.WordDocument, "M_FZGZXX1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FZGZXX2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FZGZXX3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FZGZXX4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FZGZXX5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L52", out success), "{0:F2}"); //小数点后两位
                        //5.两位小数的百分数（80KV重复性）
                        wu.WriteDataValue(wu.WordDocument, "M_BSKVCFX", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B28", out success), "{0:0.00%}"); //两位小数的百分数
                        //6.普通复制（备注说明），剂量CT在B27，KV在B29
                        if (excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B29", out success) == null || excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B29", out success).ToString() == @"/")
                        {
                            wu.WriteValue(wu.WordDocument, "M_BZSM", "无");
                        }
                        else
                        {
                            wu.WriteValue(wu.WordDocument, "M_BZSM", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B29", out success));
                        }
                        break;
                    default:
                        Log.LogHelper.AddException("生成证书时指定了不存在的检定类型", true);
                        break;
                }

                /// <summary>
                /// 类型1：普通复制
                /// </summary>
                wu.WriteValue(wu.WordDocument, "M_NAME", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B4", out success));
                wu.WriteValue(wu.WordDocument, "M_SERIAL", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F5", out success));
                wu.WriteValue(wu.WordDocument, "M_PRODUCT", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J5", out success));
                wu.WriteValue(wu.WordDocument, "M_QIJU", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "B5", out success));
                wu.WriteValue(wu.WordDocument, "M_LNGCH", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D12", out success));
                wu.WriteValue(wu.WordDocument, "M_DATE", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "K31", out success));
                wu.WriteValue(wu.WordDocument, "M_STRESS", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F4", out success));
                wu.WriteValue(wu.WordDocument, "M_QIYA", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J8", out success));
                wu.WriteValue(wu.WordDocument, "M_TEMP", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "K7", out success));
                wu.WriteValue(wu.WordDocument, "M_ZHSH", stemp1);
                wu.WriteValue(wu.WordDocument, "M_ZHSH2", stemp1);
                wu.WriteValue(wu.WordDocument, "M_ZHSH3", stemp1);

                /// <summary>
                /// 类型2：百分比换算后复制
                /// </summary>
                otemp1 = (object)excel.GetRange(excel.ExcelWorkbook, sourceIndex, "M7", out success).Value;
                if (otemp1 == null)
                {
                    stemp1 = "/";
                }
                else
                {
                    stemp1 = string.Format("{0:F1}", float.Parse(otemp1.ToString()) * 100);
                }
                wu.WriteValue(wu.WordDocument, "M_SHIDU", stemp1);
                /// <summary>
                /// 类型3：仪器编号两段合并后复制
                /// </summary>
                otemp1 = (object)excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H5", out success).Value;
                if (otemp1 == null)
                {
                    stemp1 = "";
                }
                else
                {
                    stemp1 = otemp1.ToString();
                }
                otemp1 = (object)excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L5", out success).Value;
                if (otemp1 == null)
                {
                    if (stemp1 == "")
                    {
                        stemp1 = "/";
                    }
                }
                else if (otemp1.ToString() == "/")
                {
                    if (stemp1 == "")
                    {
                        stemp1 = "/";
                    }
                }
                else
                {
                    stemp1 = stemp1 + " + " + otemp1.ToString();
                }
                wu.WriteValue(wu.WordDocument, "M_NUM", stemp1);

                //另存word和pdf
                wu.WordDocument.SaveAs2(Path.Combine(savePath, DataUtility.DataUtility.FileNameCleanName(wdName)));
                excel.SaveAsPDF(excel.ExcelWorkbook, Path.Combine(pdfPath, DataUtility.DataUtility.FileNameCleanName(pdfName)), tempFolder, out success);
                wu.WordDocument.Saved = true;
                wu.TryClose();
            }
            catch (Exception ex)
            {
                success = false;
                Log.LogHelper.AddException("生成证书时出现错误：" + ex.Message, true);
            }
        }

        public static void ReportSuspiciousFiles(Dictionary<string, int> suspFiles, string filename)
        {
            Log.LogHelper.AddException("发现" + suspFiles.Count + "个可疑匹配项，暂不作任何处理", true);
            Log.LogHelper.AddLog("可疑15", "    原文件名：" + filename, true);
            foreach (KeyValuePair<string, int> item in suspFiles)
            {
                Log.LogHelper.AddLog("可疑16", "      文件名：" + item.Key, true);
                Log.LogHelper.AddLog("可疑17", "        可疑指数：" + item.Value + " %", true);
            }
        }

    }
}
