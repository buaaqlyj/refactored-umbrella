using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;

using Statistics.Criterion;

namespace Statistics.Job
{
    public class Job
    {
        #region
        public void AddException(string message, bool save)
        {

        }
        #endregion

        #region Operation

        enum MatchPattern { FullyMatch = 2, GeneralMatch = 1, DoNotNecessary = -1, DoNotMatch = 0 };

        enum Flag { Neg = -1, Zero = 0, Pos = 1 };
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
        public FileInfo SearchForFile(string fullName, string strType, string strMacSerial, string strSensorSerial, FileInfo[] fis, out bool contin)
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
                    if (InstrumentType.IsEqualTo(type, strType))
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
                AddException("发现" + suspFiles.Count + "个可疑匹配项，暂不作任何处理", true);
                AddLog("可疑15", "    原文件名：" + fullName, true);
                foreach (KeyValuePair<string, int> item in suspFiles)
                {
                    AddLog("可疑16", "      文件名：" + item.Key, true);
                    AddLog("可疑17", "        可疑指数：" + item.Value + " %", true);
                }
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

        public FileInfo[] GetFilesFromType(string pathBase, string type, string extension, out bool checkClear)
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
                    AddException("无法识别的仪器类型：" + item, true);
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
        public void CopyData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int pattern, string certIdori, bool needFix, bool shouldFix, int startDestiRowIndex, out int newSheetIndex, out bool success)
        {
            bool success1;
            bool noNeed = false;
            int templateIndex = -1;
            int startSourceRowIndex = -1;
            int destiIndex = -1;
            MSExcel.Range rr = null;
            MSExcel.Worksheet ws1 = null;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();

            string temp;
            string text = "";
            string certId = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, "L2", out success1);
            if (!success1)
            {
                AddException(@"无法提取到证书编号", true);
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
                    AddDataError(@"第" + item.Index + "页发现多余的标准模板", true);
                }
                else
                {
                    temp = destiEx.GetText(destiEx.ExcelWorkbook, item.Index, "L2", out success1).Trim();
                    if (temp.StartsWith(@"20") && (temp.Length == 9 || temp.Length == 10))
                    {
                        //找到有证书编号的数据页
                        if (item.Name == certId)
                        {
                            if (MessageBox.Show("在历史数据记录的Excel中已发现了证书编号为" + certId + "的页面，是否覆盖？选择是，进行覆盖。选择否，停止对本Excel的处理", "是否覆盖", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                            {
                                ws1 = item;
                            }
                            else
                            {
                                AddException(@"要合并入的数据已存在于第" + item.Index + "页", true);
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
                    AddException(@"找不到数据的标准模板", true);
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
                            AddException(@"标准模板复制出错", true);
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
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i, 3, out success1).Trim();
                        if (text == "1")
                        {
                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i + 1, 3, out success1).Trim();
                            if (text == "2")
                            {
                                text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, i + 2, 3, out success1).Trim();
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
                        AddException(@"找不到原始数据所在的行", true);
                        success = false;
                        newSheetIndex = -1;
                        return;
                    }

                    //拷贝数据
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 2, 12, "@", out success1);
                    if (!success1) AddException(@"《证书编号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 4, 1, 4, 2, new string[] { @"送校单位：", @"单位名称：" }, "@", out success1);
                    if (!success1) AddException(@"《送校单位》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 4, 5, 4, 6, @"联系地址：", "@", out success1);
                    //if (!success1) AddException(@"《联系地址》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 1, 5, 2, @"仪器名称：", "@", out success1);
                    if (!success1) AddException(@"《仪器名称》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 5, 5, 6, @"型号：", "@", out success1);
                    if (!success1) AddException(@"《型号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 7, 5, 8, new string[] { @"主机编号：", @"编号：" }, "@", out success1);
                    if (!success1) AddException(@"《主机编号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 9, 5, 10, @"厂家：", "@", out success1);
                    if (!success1) AddException(@"《厂家》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 5, 11, 5, 12, new string[] { @"探测器编号：", "电离室号：", "探测器号：" }, "@", out success1);
                    if (!success1) AddException(@"《探测器编号》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 31, 7, "", out success1);
                    //if (!success1) AddException(@"《记录者》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 31, 9, "", out success1);
                    //if (!success1) AddException(@"《校对者》数据复制错误", true);
                    CopyDate(sourceEx, sourceIndex, destiEx, ws1.Index, out success1);

                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 7, 11, "0.000", out success1);
                    if (needFix && !success1) AddException(@"《温度》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 7, 13, "0.0%", out success1);
                    if (needFix && !success1) AddException(@"《湿度》数据复制错误", true);
                    CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 8, 10, "", out success1);
                    if (needFix && !success1) AddException(@"《气压》数据复制错误", true);

                    rr = destiEx.GetRange(destiEx.ExcelWorkbook, ws1.Index, "M8", out success1);
                    if (needFix)
                    {
                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 13, "修正", out success1);
                    }
                    else
                    {
                        //电离室->半导体
                        rr = destiEx.GetRange(destiEx.ExcelWorkbook, ws1.Index, "L8", out success1);
                        rr.FormulaLocal = "";
                        rr.Formula = "";
                        rr.FormulaArray = "";
                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 12, "1.000000", "@", out success1);
                        if (shouldFix)
                        {
                            destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 13, "自修正", out success1);
                        }
                        else
                        {
                            destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 8, 13, "不修正", out success1);
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
                        text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, p, 1, out success1).Trim() + sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, p, 2, out success1).Trim() + sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, p, 3, out success1).Trim();
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
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 27, 13, "@", out success1);
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, "", out success1);
                            if (!success1) AddException(@"《量程》数据复制错误", true);

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out success1).Trim();
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 4, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 6, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 8, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 10, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeDoseData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 12, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 12, out success1).Trim();
                            if (text.StartsWith("单位"))
                            {
                                CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 13, "", out success1);
                                if (!success1) AddException(@"《单位》数据复制错误", true);
                            }
                            else
                            {
                                dataStruct = DataStruct.CalDataRange(dataStructList, sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out success1).Trim(), pattern, out success);
                                switch (dataStruct.DataRanges)
                                {
                                    case DataRange.cGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "cGy", out success1);
                                        break;
                                    case DataRange.mGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mGy", out success1);
                                        break;
                                    case DataRange.mR:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mR", out success1);
                                        break;
                                    case DataRange.R:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "R", out success1);
                                        break;
                                    case DataRange.uGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "μGy", out success1);
                                        break;
                                    case DataRange.Unknown:
                                        AddException("无法判断数据单位", true);
                                        break;
                                }
                                switch (dataStruct.Distance)
                                {
                                    case Distance.d1:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.0m", out success1);
                                        break;
                                    case Distance.d1_5:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.5m", out success1);
                                        break;
                                    case Distance.Unknown:
                                        AddException("无法获取标准值，判断测试距离", true);
                                        break;
                                }
                            }

                            break;
                        case 1:
                            //CT
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 27, 13, "@", out success1);
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, "", out success1);
                            if (!success1) AddException(@"《量程》数据复制错误", true);

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out success1).Trim();
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 4, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 6, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 8, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 10, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));
                            dataStructList.Add(CopyThreeCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 12, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1));

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 12, out success1).Trim();
                            if (text.StartsWith("单位"))
                            {
                                CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 13, "", out success1);
                                if (!success1) AddException(@"《单位》数据复制错误", true);
                            }
                            else
                            {
                                dataStruct = DataStruct.CalDataRange(dataStructList, sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out success1).Trim(), pattern, out success);
                                switch (dataStruct.DataRanges)
                                {
                                    case DataRange.mGycm:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mGycm", out success1);
                                        break;
                                    case DataRange.mGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mGy", out success1);
                                        break;
                                    case DataRange.cGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "cGy", out success1);
                                        break;
                                    case DataRange.mR:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "mR", out success1);
                                        break;
                                    case DataRange.R:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "R", out success1);
                                        break;
                                    case DataRange.uGy:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "μGy", out success1);
                                        break;
                                    case DataRange.cGycm:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "cGycm", out success1);
                                        break;
                                    case DataRange.Rcm:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 12, 13, "Rcm", out success1);
                                        break;
                                    case DataRange.Unknown:
                                        AddException("无法判断数据单位", true);
                                        break;
                                }
                                switch (dataStruct.Distance)
                                {
                                    case Distance.d1:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.0m", out success1);
                                        break;
                                    case Distance.d1_5:
                                        destiEx.WriteValue(destiEx.ExcelWorkbook, ws1.Index, 6, 13, "1.5m", out success1);
                                        break;
                                    case Distance.Unknown:
                                        AddException("无法获取标准值，判断测试距离", true);
                                        break;
                                }
                            }
                            //CopyCTData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, startDestiRowIndex, "", out success1);
                            break;
                        case 2:
                            //KV
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 28, 13, "@", out success1);
                            //if (!success1) AddException(@"《备注》数据复制错误", true);
                            CopyOneData(sourceEx, sourceIndex, destiEx, ws1.Index, 12, 4, @"/", "", out success1);
                            if (!success1) AddException(@"《量程》数据复制错误", true);

                            text = sourceEx.GetText(sourceEx.ExcelWorkbook, sourceIndex, 12, 4, out success1).Trim();
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 4, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 6, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 8, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 10, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1);
                            CopyThreeKVData(sourceEx, sourceIndex, destiEx, ws1.Index, startSourceRowIndex, 12, startDestiRowIndex, "", standardRowIndex, text, pattern, out success1);

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

        public void CopyData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int pattern, string certIdori, bool needFix, bool shouldFix, int startDestiRowIndex, out bool success)
        {
            int index = 0;
            CopyData(sourceEx, sourceIndex, destiEx, pattern, certIdori, needFix, shouldFix, startDestiRowIndex, out index, out success);
        }

        public void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, string style, out bool sc)
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

        public void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, string style, bool checkDouble, out bool sc)
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

        public void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, string defaultValue, string style, out bool sc)
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

        public void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, int new_row, int new_column, string pre, string style, out bool sc)
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

        public void CopyOneData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int rowIndex, int columnIndex, int new_row, int new_column, string[] pre, string style, out bool sc)
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

        public void CopyThreeKVData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int startSourceRowIndex, int columnIndex, int startDestiRowIndex, string style, int standardRowIndex, string range, int pattern, out bool sc)
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
                        AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
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
                        AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
                        sc = false;
                    }
                }
            }
        }

        public DataStruct CopyThreeDoseData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int startSourceRowIndex, int columnIndex, int startDestiRowIndex, string style, int standardRowIndex, string range, int pattern, out bool sc)
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
                        AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
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
                        AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
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

        public DataStruct CopyThreeCTData(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, int startSourceRowIndex, int columnIndex, int startDestiRowIndex, string style, int standardRowIndex, string range, int pattern, out bool sc)
        {
            bool checkClear;
            bool conti = true;
            double temp_double = 0.0;
            double diValue = 0.00001;
            string text;
            MSExcel.Range rr = null;
            DoseCriterion dataCri = DoseCriterion.Null;

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
                                AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
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
                                AddException(@"第" + columnIndex + "列第" + (startSourceRowIndex + i).ToString() + "行不包含有效数据", true);
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
                    AddException("规范内容有误：" + text, true);
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
        public bool GetFixState(ExcelUtility sourceEx, int sourceIndex, int pattern, out bool needFix, out bool shouldFix)
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
                AddException("无法判断电离室与半导体类型", true);
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
        public bool TestFlag(ExcelUtility sourceEx, int sourceIndex, int startSourceRowIndex, int columnIndex)
        {
            Flag a, b, c;
            bool conti = false;
            double dig;
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
                    AddException("数据记录有未识别的数据", true);
                }
            }
            return conti;
        }

        public void CopyDate(ExcelUtility sourceEx, int sourceIndex, ExcelUtility destiEx, int destiIndex, out bool sc)
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
                AddException(@"《日期》数据复制错误：" + text, true);
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
        public bool Statistic(ExcelUtility eu, int pattern)
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
            Dictionary<DoseCriterion, int> existCriList = new Dictionary<DoseCriterion, int>();

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
        public bool Statistic(ExcelUtility eu, int pattern, bool Perfect, string strCompany, string strType, string tempName, string certstr)
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
            Dictionary<DoseCriterion, int> existDoseCriList = new Dictionary<DoseCriterion, int>() { 
                { DoseCriterion.RQR2_40  , DoseCriterion.RQR2_40.Column  }, 
                { DoseCriterion.RQR3_50  , DoseCriterion.RQR3_50.Column  },
                { DoseCriterion.RQR4_60  , DoseCriterion.RQR4_60.Column  },
                { DoseCriterion.RQR5_70  , DoseCriterion.RQR5_70.Column  },
                { DoseCriterion.RQR6_80  , DoseCriterion.RQR6_80.Column  },
                { DoseCriterion.RQR7_90  , DoseCriterion.RQR7_90.Column  },
                { DoseCriterion.RQR8_100 , DoseCriterion.RQR8_100.Column },
                { DoseCriterion.RQR9_120 , DoseCriterion.RQR9_120.Column },
                { DoseCriterion.RQR_140  , DoseCriterion.RQR_140.Column  },
                { DoseCriterion.RQR10_150, DoseCriterion.RQR10_150.Column}
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
        public void TypeStandardize(ExcelUtility sourceEx, int stateIndex)
        {
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
            foreach (string item in InstrumentType.AllTypes)
            {
                if (item.ToLower() == str)
                {
                    DoNotHave = false;
                    break;
                }
            }
            if (DoNotHave)
            {
                AddException(@"仪器类型可能出现手误", true);
                AddLog("错误47", "  仪器类型：" + str, true);
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
        public void GenerateCert(ExcelUtility excel, int sourceIndex, int pattern, string wordPath, string savePath, string pdfPath, string tempFolder, bool shouldFix, out bool success)
        {
            //GenerateCert(_sr, stateIndex, path, pS.CertFolder, pS.PDFDataFolder, out success);
            try
            {
                WordUtility wu = new WordUtility(wordPath, out success);
                if (!success)
                {
                    AddException("Word文档打开失败", true);
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
                        wu.WriteDataValue(wu.WordDocument, "M_FFGZXX1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "D52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FFGZXX2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "F52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FFGZXX3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "H52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FFGZXX4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "J52", out success), "{0:F2}"); //小数点后两位
                        wu.WriteDataValue(wu.WordDocument, "M_FFGZXX5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, "L52", out success), "{0:F2}"); //小数点后两位
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
                        AddException("生成证书时指定了不存在的检定类型", true);
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
                AddException("生成证书时出现错误：" + ex.Message, true);
            }
        }

        //public void DoWork(JobParameterStruct pS)
        //{
        //    string inputFi = pS.InputFile;
        //    string inputFo = pS.AutoInputFolder;
        //    string outputFo = pS.AutoOutputFolder;
        //    string inputEx = pS.AutoExtension;
        //    int cbsi = pS.DataPattern;

        //    if (pS.ActionType < 0 || pS.ActionType > 4)
        //    {
        //        if (MessageBox.Show(@"没有选择任何操作类型，请点击确定后返回重试", "提示", MessageBoxButtons.OK) == DialogResult.OK) { }
        //    }
        //    else if (pS.ActionType < 1 && (pS.FixType < 0 || pS.FixType > 1))
        //    {
        //        if (MessageBox.Show(@"没有选择温度气压修正选项，请点击确定后返回重试", "提示", MessageBoxButtons.OK) == DialogResult.OK) { }
        //    }
        //    else if (MessageBox.Show(@"处理成功将会删除原纪录，是否要继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //    {
        //        List<string> probfile = new List<string>();
        //        UpdateProgress(0);

        //        if (inputFi != "")
        //        {
        //            if (!inputFi.Contains(@"~$"))
        //            {
        //                currentFile = inputFi;
        //                ToolStripStatusLabel_SetText(new FileInfo(inputFi).Name);
        //                try
        //                {
        //                    dataerrorNum = 0;
        //                    exceptionNum = 0;
        //                    dod(inputFi, pS, ref probfile);
        //                }
        //                catch (Exception ex)
        //                {
        //                    probfile.Add(inputFi);
        //                    if (exceptionNum > 0)
        //                    {
        //                        exceptionNum = 0;
        //                    }
        //                    else
        //                    {
        //                        AddLog(@"***************************************************************", true);
        //                    }
        //                    AddLog(@"* 文件名称：" + inputFi, true);
        //                    AddLog(@"* 异常消息：" + ex.Message, true);
        //                    AddLog(@"* 异常方法：" + ex.TargetSite, true);
        //                    AddLog(@"***************************************************************", true);
        //                }
        //                if (dataerrorNum > 0)
        //                {
        //                    dataerrorNum = 0;
        //                    AddLog(@"***************************************************************", true);
        //                }
        //                ToolStripStatusLabel_SetText("");
        //            }

        //            UpdateProgress(100);

        //            if (probfile.Count > 0)
        //            {
        //                AddLog(@"结果02", "共处理文件数：1", true);
        //                AddLog(@"结果03", "异常文件总数：" + probfile.Count, true);
        //                foreach (string item in probfile)
        //                {
        //                    AddLog(@"结果04", item, true);
        //                }
        //            }
        //            else
        //            {
        //                AddLog(@"结果05", "没有发现任何异常，共处理文件数：1", true);
        //            }
        //            MessageBox.Show("共处理文件" + "1" + "个。" + Environment.NewLine + "其中成功" + (1 - probfile.Count).ToString() + "个，失败" + probfile.Count + "个。", "处理完成", MessageBoxButtons.OK);
        //        }
        //        else if (inputFo != "")
        //        {
        //            DirectoryInfo di = new DirectoryInfo(inputFo);
        //            FileInfo[] fis = di.GetFiles(inputEx, SearchOption.AllDirectories);
        //            probfile.Clear();
        //            int doneNumber = 0;
        //            if (fis.Length > 0)
        //            {
        //                double Step = 100.0 / (double)fis.Length;
        //                double values = 0;
        //                foreach (FileInfo item in fis)
        //                {
        //                    if (!item.Name.StartsWith(@"~$"))
        //                    {
        //                        currentFile = item.FullName;
        //                        ToolStripStatusLabel_SetText(new FileInfo(item.FullName).Name);
        //                        try
        //                        {
        //                            dataerrorNum = 0;
        //                            exceptionNum = 0;
        //                            dod(item.FullName, pS, ref probfile);
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            probfile.Add(item.FullName);
        //                            if (exceptionNum > 0)
        //                            {
        //                                exceptionNum = 0;
        //                            }
        //                            else
        //                            {
        //                                AddLog(@"***************************************************************", true);
        //                            }
        //                            AddLog(@"* 文件名称：" + item.FullName, true);
        //                            AddLog(@"* 异常消息：" + ex.Message, true);
        //                            AddLog(@"* 异常方法：" + ex.TargetSite, true);
        //                            AddLog(@"***************************************************************", true);
        //                        }
        //                        if (dataerrorNum > 0)
        //                        {
        //                            dataerrorNum = 0;
        //                            AddLog(@"***************************************************************", true);
        //                        }
        //                        ToolStripStatusLabel_SetText("");
        //                    }
        //                    values += Step;
        //                    UpdateProgress(values);
        //                    doneNumber++;
        //                    if (isStopping)
        //                    {
        //                        break;
        //                    }
        //                }
        //                if (probfile.Count > 0)
        //                {
        //                    AddLog(@"结果06", "异常文件总数：" + probfile.Count, true);
        //                    foreach (string item in probfile)
        //                    {
        //                        AddLog(@"结果07", item, true);
        //                    }
        //                }
        //                else
        //                {
        //                    AddLog(@"结果08", "没有发现任何异常，共处理文件数：" + doneNumber, true);
        //                }
        //                MessageBox.Show("共处理文件" + fis.Length + "个。" + Environment.NewLine + "其中成功" + (fis.Length - probfile.Count).ToString() + "个，失败" + probfile.Count + "个。", "处理完成", MessageBoxButtons.OK);
        //            }
        //            else
        //            {
        //                MessageBox.Show(@"输入文件夹没有找到待处理的文件");
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show(@"未指定输入文件");
        //        }
        //    }
        //    isWorking = false;
        //    isStopping = false;
        //    Button_SetText(@"开始");
        //    AddLog(Environment.NewLine + Environment.NewLine, true);
        //}

        ///// <summary>
        ///// pattern 1 KV
        ///// pattern 2 Dose
        ///// pattern 3 CT
        ///// pattern 4
        ///// </summary>
        ///// <param name="filePath"></param>
        ///// <param name="pattern"></param>
        //public void StatisticsOne(string filePath, int pattern, string text1, string text7, string text3, ref List<string> problemFilesList)
        //{
        //    Dictionary<int, string> exSheets = new Dictionary<int, string>();
        //    Dictionary<int, string> cert = new Dictionary<int, string>();
        //    List<string> sheetsName = new List<string>();
        //    bool Perfect;
        //    int stateIndex = -1;
        //    FileInfo temp_fi = null;
        //    int lineIndex = 8;
        //    int countIndex = 0;
        //    bool firstTime = true;
        //    bool needInitialState = true;
        //    string tempName;
        //    string insNumber = "";
        //    string standardInsNumber = "";
        //    string sourceDirectory = "";//, temp_text3 = "";
        //    string strCompany = "", strType = "", strMacSerial = "", strSensorSerial = "";
        //    string temp_strCompany = "", temp_strType = "", temp_strMacSerial = "", temp_strSensorSerial = "";

        //    ExcelUtility _eu = new ExcelUtility(filePath, out checkClear);
        //    if (!checkClear) { return; }
        //    _eu.ExcelApp.DisplayAlerts = false;
        //    _eu.ExcelApp.AlertBeforeOverwriting = false;
        //    Object format = MSExcel.XlFileFormat.xlWorkbookDefault;
        //    FileInfo fi = new FileInfo(_eu.path);
        //    //输入文件的目录
        //    if (text1 == "")
        //    {
        //        sourceDirectory = fi.DirectoryName.Replace(text7, text3) + @"\";
        //    }
        //    else
        //    {
        //        sourceDirectory = text3 + @"\";
        //    }

        //    try
        //    {
        //        //1.初始化：规范统计sheet结构，写入固定内容
        //        stateIndex = _eu.InitialStatisticSheet(out exSheets, out needInitialState, out checkClear);
        //        if (!checkClear) { return; }
        //        _eu.WriteTitle(pattern, stateIndex);

        //        switch (pattern)
        //        {
        //            case 0:
        //                //2.搬运数据
        //                countIndex = 0;
        //                foreach (int item in exSheets.Keys)
        //                {
        //                    countIndex++;
        //                    cert.Add(lineIndex, exSheets[item]);
        //                    _eu.CopyOneYearData(item, stateIndex, lineIndex, exSheets[item], out insNumber);
        //                    tempName = _eu.PerfectFileName(item, out temp_strCompany, out temp_strType, out temp_strMacSerial, out temp_strSensorSerial, out Perfect);
        //                    if (temp_strCompany.Length > strCompany.Length)
        //                    {
        //                        strCompany = temp_strCompany;
        //                    }
        //                    if (temp_strMacSerial.Length > strMacSerial.Length)
        //                    {
        //                        strMacSerial = temp_strMacSerial;
        //                    }
        //                    if (temp_strType.Length > strType.Length)
        //                    {
        //                        strType = temp_strType;
        //                    }
        //                    if (temp_strSensorSerial.Length > strSensorSerial.Length)
        //                    {
        //                        strSensorSerial = temp_strSensorSerial;
        //                    }
        //                    if (countIndex == 1)
        //                    {
        //                        standardInsNumber = insNumber;
        //                    }
        //                    else
        //                    {
        //                        if (GetPureNumber(standardInsNumber, true) != GetPureNumber(insNumber, true))
        //                        {
        //                            AddException(cert[8] + "和" + cert[lineIndex] + "的仪器编号不同：", true);
        //                            AddLog(@"错误", "    " + cert[8] + @"：" + standardInsNumber, true);
        //                            AddLog(@"错误", "    " + cert[lineIndex] + @"：" + insNumber, true);
        //                        }
        //                    }
        //                    if (firstTime && Perfect)
        //                    {
        //                        _eu.DesiredName = CleanName(tempName);
        //                        //写入送校单位
        //                        _eu.WriteValue(stateIndex, 2, 2, strCompany, out checkClear);
        //                        //写入仪器名称
        //                        _eu.WriteValue(stateIndex, 3, 2, strType, out checkClear);
        //                        firstTime = false;
        //                    }
        //                    lineIndex++;
        //                }

        //                if (firstTime)
        //                {
        //                    //写入送校单位
        //                    _eu.WriteValue(stateIndex, 2, 2, strCompany, out checkClear);
        //                    //写入仪器名称
        //                    _eu.WriteValue(stateIndex, 3, 2, strType, out checkClear);
        //                    //文件名
        //                    _eu.DesiredName = CleanName(strCompany + "_" + strType + "_" + strMacSerial + ((strSensorSerial.Replace(@"/", "").Trim() != "") ? "+" + strSensorSerial : ""));
        //                }
        //                //3.分析数据

        //                if (exSheets.Count > 1)
        //                {
        //                    _eu.WriteValue(stateIndex, lineIndex, 1, @"重复性", out checkClear);
        //                    //计算平均值和年稳定性
        //                    _eu.StatisticsOneColumn(stateIndex, lineIndex, 4, cert, "RQR4（60kV）", exSheets.Count > 1);
        //                    _eu.StatisticsOneColumn(stateIndex, lineIndex, 6, cert, "RQR5（70kV）", exSheets.Count > 1);
        //                    _eu.StatisticsOneColumn(stateIndex, lineIndex, 8, cert, "RQR6（80kV）", exSheets.Count > 1);
        //                    _eu.StatisticsOneColumn(stateIndex, lineIndex, 10, cert, "RQR8（100kV）", exSheets.Count > 1);
        //                    _eu.StatisticsOneColumn(stateIndex, lineIndex, 12, cert, "RQR9（120kV）", exSheets.Count > 1);
        //                }

        //                if (exceptionNum > 0)
        //                {
        //                    problemFilesList.Add(_eu.DesiredName);
        //                }
        //                break;
        //            case 1:

        //                break;
        //        }

        //        temp_fi = new FileInfo(RightFileName(sourceDirectory, _eu.DesiredName, fi.Extension, "-new"));
        //        if (!temp_fi.Directory.Exists)
        //        {
        //            Directory.CreateDirectory(temp_fi.DirectoryName);
        //        }
        //        _eu.ExcelDoc.SaveAs(temp_fi.FullName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, MSExcel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //    }
        //    catch (Exception ex)
        //    {
        //        exceptionNum++;
        //        if (exceptionNum == 1)
        //        {
        //            AddLog(@"信息", "原文件名：" + filePath, true);
        //        }
        //        AddLog(@"异常", ex.Message, true);
        //        AddLog(@"异常", "  " + ex.TargetSite.ToString(), true);
        //    }
        //    finally
        //    {
        //        //移动前的输出文件路径
        //        sourceDirectory = sourceDirectory + _eu.DesiredName + fi.Extension;

        //        _eu.ExcelDoc.Saved = true;
        //        _eu.TryClose();
        //        if (exceptionNum > 0)
        //        {

        //            //移动后的输出文件路径
        //            tempName = RightFileName(Application.StartupPath + @"\有问题整理后\", _eu.DesiredName, fi.Extension, "-new");

        //            AddLog(@"信息", "新文件名：" + tempName, true);
        //            AddLog(@"***************************************************************", true);

        //            if (File.Exists(sourceDirectory)) File.Move(sourceDirectory, tempName);

        //            if (text1 == "")
        //            {
        //                temp_fi = new FileInfo(fi.FullName.Replace(text7, Application.StartupPath + @"\有问题"));
        //                if (!temp_fi.Directory.Exists)
        //                {
        //                    Directory.CreateDirectory(temp_fi.DirectoryName);
        //                }
        //                File.Copy(fi.FullName, temp_fi.FullName, true);
        //            }
        //            else
        //            {
        //                File.Copy(fi.FullName, Application.StartupPath + @"\有问题\" + fi.Name + fi.Extension, true);
        //            }
        //        }
        //    }
        //}

        #endregion

        #region TestingFunction
        #region GeneratingForm

        public void GeneratingForm(string filePath, JobParameterStruct pS, ref List<string> problemFilesList)
        {
            //int pattern = pS.DataPattern;
            string output = pS.AutoOutputFolder;
            bool needFix = pS.AutoFixType;
            string templateName = pS.DataTemplateFilePath;
            string macType = pS.MacType;

            bool success = true;

            int templateIndex = -1;

            MSExcel.Worksheet ws1 = null;

            WordUtility _wu = new WordUtility(filePath, out success);
            if (!success)
            {
                AddException("Word文档打开失败", true);
                return;
            }

            string tempName = _wu.GetText(_wu.WordDocument, 7);//B4:送校单位
            string tempSerial = _wu.GetText(_wu.WordDocument, 15).Trim();//F5:仪器型号
            //string tempNum = _wu.GetText(_wu.WordDocument, 19);//*H5:仪器编号
            string tempQiju = _wu.GetText(_wu.WordDocument, 11);//B5:仪器名称
            string tempZhsh = _wu.GetText(_wu.WordDocument, 3);//L2:证书编号
            string tempStress = _wu.GetText(_wu.WordDocument, 27);//F4:联系地址

            _wu.TryClose();

            if (tempSerial != "" && tempSerial != macType)
            {
                AddDataError("证书中包含的仪器型号与指定的仪器型号不符", true);
            }

            string str = tempZhsh.Substring(8);
            string strSavename = Path.Combine(output, tempName + "_" + macType + "_" + str + ".xlsx");

            if (File.Exists(strSavename))
            {
                if (MessageBox.Show(@"文件已存在，是否覆盖？" + Environment.NewLine + strSavename, "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    File.Delete(strSavename);
                }
                else
                {
                    success = false;
                    return;
                }
            }
            File.Copy(templateName, strSavename);

            ExcelUtility _sr = new ExcelUtility(strSavename, out checkClear);
            if (!checkClear)
            {
                AddException(@"Excel文档无法打开", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            _sr.ExcelApp.DisplayAlerts = false;
            _sr.ExcelApp.AlertBeforeOverwriting = false;

            try
            {
                foreach (MSExcel.Worksheet item in _sr.ExcelWorkbook.Sheets)
                {
                    if (item.Name == @"标准模板")
                    {
                        templateIndex = item.Index;
                    }
                    else if (item.Name.Contains(@"标准模板"))
                    {
                        AddException(@"发现多余的标准模板", true);
                    }
                }
                if (templateIndex > -1)
                {
                    ws1 = (MSExcel.Worksheet)_sr.ExcelWorkbook.Sheets[templateIndex];
                    ws1.Copy(ws1, Type.Missing);
                    ws1 = (MSExcel.Worksheet)_sr.ExcelWorkbook.Sheets[templateIndex];
                    if (!ws1.Name.Contains(@"标准模板"))
                    {
                        AddException(@"标准模板复制出错", true);
                        success = false;
                        return;
                    }
                    else
                    {
                        ws1.Name = str;
                    }
                }
                else
                {
                    AddException(@"找不到模板excel中的标准模板页", true);
                }

                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 4, 2, tempName, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 5, 6, macType, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 5, 2, tempQiju, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 2, 12, str, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 4, 6, tempStress, out success);

                if (needFix)
                {
                    _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 8, 13, "修正", out success);
                }
                else
                {
                    //电离室->半导体
                    MSExcel.Range rr = _sr.GetRange(_sr.ExcelWorkbook, ws1.Index, "L8", out success);
                    rr.FormulaLocal = "";
                    rr.Formula = "";
                    rr.FormulaArray = "";
                    _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 8, 12, "1.000000", "@", out success);
                    _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 8, 13, "不修正", out success);
                }
                //写入记录者
                _sr.WriteImage(_sr.ExcelWorkbook, ws1.Index, 29, 7, person, 45, 28, out success);

                _sr.ExcelWorkbook.Save();
            }
            catch (Exception ex)
            {
                AddException("生成证书时遇到异常：" + ex.Message, true);
            }
            finally
            {
                //关闭Excel
                if (_sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                //有重大失误的情况下报错，没有失误就删除源word文件
                if (exceptionNum > 0)
                {
                    AddLog(@"***************************************************************", true);
                    problemFilesList.Add(filePath);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                }
                else
                {
                    File.Delete(filePath);
                }
            }
        }

        #endregion

        #region GeneratingCertificate

        public void GeneratingCertificateOne(string filePath, JobParameterStruct pS, ref List<string> problemFilesList)
        {
            int pattern = pS.DataPattern;
            string output = pS.AutoOutputFolder;
            string ext = pS.AutoExtension;
            bool createNew = pS.CreateNew;
            string tempFo = pS.TempFolder;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();

            bool Perfect;
            bool hasArchieved = false;
            bool needFix = true;
            bool shouldFix = true;
            bool needTestGeCe = true, canGeCe = false;

            int stateIndex = -1;
            int noIdNumber = 0;
            int startDestiRowIndex = -1;

            string tempName;
            string certId;
            string newName = "";
            string fileText = "";
            string strCompany = "", strType = "", strMacSerial = "", strSensorSerial = "";

            string backupKey = "";

            FileInfo temp_fi = null;
            FileInfo fi = new FileInfo(filePath);

            Object format = MSExcel.XlFileFormat.xlWorkbookDefault;

            ExcelUtility _sr = new ExcelUtility(filePath, out checkClear);
            ExcelUtility _eu = null;
            MSExcel.Range rr = null;
            if (!checkClear)
            {
                AddException(@"Excel文档无法打开", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            _sr.ExcelApp.DisplayAlerts = false;
            _sr.ExcelApp.AlertBeforeOverwriting = false;

            //第一次循环：获取信息，并规范每页的标签
            foreach (MSExcel.Worksheet item in _sr.ExcelWorkbook.Sheets)
            {
                //规范sheet标签名为证书编号
                certId = _sr.GetText(_sr.ExcelWorkbook, item.Index, "L2", out checkClear).Trim();
                if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                {
                    //有规范的证书号
                    exSheets.Add(item.Index, certId);
                    stateIndex = item.Index;
                }
                else
                {
                    //无规范的证书号
                    rr = _sr.GetRange(_sr.ExcelWorkbook, item.Index, "A4", out checkClear);
                    if (!item.Name.Contains(@"标准模板") && rr.Text.ToString().Trim().StartsWith(@"送校单位"))
                    {
                        //有记录不包含规范的证书编号
                        AddException(@"该文档有实验数据不包含证书编号", true);
                        noIdNumber++;
                    }
                }
            }

            if (exSheets.Count == 0)
            {
                if (noIdNumber == 0)
                {
                    AddException(@"该文档可能是空文档", true);
                }
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            else if (exSheets.Count + noIdNumber > 1)
            {
                AddException(@"该文档包含多个数据sheet，默认处理第一个", true);
            }

            certId = exSheets[stateIndex];
            tempName = _sr.GenerateFileName(_sr.ExcelWorkbook, stateIndex, out strCompany, out strType, out strMacSerial, out strSensorSerial, out Perfect);
            checkClear = false;
            if (strCompany == "")
            {
                AddException(@"送校单位信息未提取到", true);
                checkClear = true;
            }
            if (strType == "")
            {
                AddException(@"仪器型号信息未提取到", true);
                checkClear = true;
            }
            if (strMacSerial == "")
            {
                AddException(@"主机编号信息未提取到", true);
                checkClear = true;
            }
            if (checkClear)
            {
                AddLog(@"***************************************************************", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
            //判断目标行数
            if (pattern < 2)
            {
                startDestiRowIndex = 18;
            }
            else if (pattern == 2)
            {
                //TODO: kV修正项
            }
            else
            {
                AddException(@"数据类型无效", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }

            try
            {
                //寻找目标统计文件
                existFile = GetFilesFromType(output, _sr.GetText(_sr.ExcelWorkbook, stateIndex, "F5", out checkClear), ext, out checkClear);
                if (!checkClear)
                {
                    if (_sr != null && _sr.ExcelWorkbook != null)
                    {
                        _sr.ExcelWorkbook.Saved = true;
                        _sr.TryClose();
                    }
                    AddLog(@"***************************************************************", true);
                    problemFilesList.Add(filePath);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                    return;
                }
                //output = Path.Combine(output, _sr.GetText(_sr.ExcelWorkbook, stateIndex, "F5", out checkClear));
                //if (!Directory.Exists(output))
                //{
                //    AddException("无法识别的仪器类型：" + _sr.GetText(_sr.ExcelWorkbook, stateIndex, "F5", out checkClear), true);
                //    if (_sr != null && _sr.ExcelWorkbook != null)
                //    {
                //        _sr.ExcelWorkbook.Saved = true;
                //        _sr.TryClose();
                //    }
                //    AddLog(@"***************************************************************", true);
                //    problemFilesList.Add(filePath);
                //    exceptionNum = 0;
                //    dataerrorNum = 0;
                //    return;
                //}
                //existFile = (new DirectoryInfo(output)).GetFiles(@"*.xls*", SearchOption.AllDirectories);
                temp_fi = SearchForFile(filePath, strType, strMacSerial, strSensorSerial, existFile, out hasArchieved);
                if (temp_fi == null)
                {
                    if (createNew)
                    {
                        //对新记录建档
                        needTestGeCe = false;
                        if (hasArchieved)
                        {
                            //不存在，复制当前记录过去
                            newName = DataUtility.DataUtility.PathRightFileName(output, tempName, fi.Extension, "_new");
                            _sr.ExcelWorkbook.SaveAs(newName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            temp_fi = new FileInfo(newName);
                            Statistic(_sr, pattern, Perfect, strCompany, strType, tempName, certId);
                            //TODO: 生成证书处校验是否可选（不存在，复制当前记录过去）
                            _sr.Verification(_sr.ExcelWorkbook, true, pattern, out checkClear);
                            fileText = newName;
                        }
                        else
                        {
                            //不存在，但搜索到有可疑项，暂不处理
                            fileText = filePath;
                        }
                    }
                    else
                    {
                        //对新记录不建档，退出
                        AddException("没有在历史数据记录中发现匹配项，暂不处理。", true);
                        if (_sr != null && _sr.ExcelWorkbook != null)
                        {
                            _sr.ExcelWorkbook.Saved = true;
                            _sr.TryClose();
                        }
                        AddLog(@"***************************************************************", true);
                        problemFilesList.Add(filePath);
                        exceptionNum = 0;
                        dataerrorNum = 0;
                        return;
                    }
                }
                else if (File.Exists(temp_fi.FullName))
                {
                    DataUtility.DataUtility.BackupFile(tempFo, temp_fi.FullName, out backupKey);
                    _eu = new ExcelUtility(temp_fi.FullName, out checkClear);
                    if (!checkClear)
                    {
                        AddException(@"Excel文档无法打开", true);
                        if (_sr != null && _sr.ExcelWorkbook != null)
                        {
                            _sr.ExcelWorkbook.Saved = true;
                            _sr.TryClose();
                        }
                        if (_eu != null && _eu.ExcelWorkbook != null)
                        {
                            _eu.ExcelWorkbook.Saved = true;
                            _eu.TryClose();
                        }
                        AddLog(@"***************************************************************", true);
                        problemFilesList.Add(filePath);
                        exceptionNum = 0;
                        dataerrorNum = 0;
                        return;
                    }
                    _eu.ExcelApp.DisplayAlerts = false;
                    _eu.ExcelApp.AlertBeforeOverwriting = false;

                    CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out checkClear);
                    if (!checkClear)
                    {
                        if (_sr != null && _sr.ExcelWorkbook != null)
                        {
                            _sr.ExcelWorkbook.Saved = true;
                            _sr.TryClose();
                        }
                        if (_eu != null && _eu.ExcelWorkbook != null)
                        {
                            _eu.ExcelWorkbook.Saved = true;
                            _eu.TryClose();
                        }
                        problemFilesList.Add(fileText);
                        AddLog(@"***************************************************************", true);
                        exceptionNum = 0;
                        dataerrorNum = 0;
                        return;
                    }
                    _eu.ExcelWorkbook.Save();
                    _eu.ExcelWorkbook.Saved = true;
                    canGeCe = Statistic(_eu, pattern, Perfect, strCompany, strType, tempName, certId);
                    _eu.ExcelWorkbook.Save();
                    _eu.ExcelWorkbook.Saved = true;
                    //TODO: 生成证书处校验是否可选（存在，合并入原记录）
                    _eu.Verification(_eu.ExcelWorkbook, false, pattern, out checkClear);
                    fileText = temp_fi.FullName;
                }
                else
                {
                    AddException(@"文件不存在：" + temp_fi.FullName, true);
                    fileText = filePath;
                }
            }
            catch (Exception ex)
            {
                AddException(@"生成证书时合并一步遇到异常：" + ex.Message, true);
            }

            try
            {
                //有重大失误，关闭两个excel，报错退出。没有失误继续运行
                if (exceptionNum > 0)
                {
                    if (_sr != null && _sr.ExcelWorkbook != null)
                    {
                        _sr.ExcelWorkbook.Saved = true;
                        _sr.TryClose();
                    }
                    if (_eu != null && _eu.ExcelWorkbook != null)
                    {
                        _eu.ExcelWorkbook.Saved = true;
                        _eu.TryClose();
                    }
                    problemFilesList.Add(fileText);
                    AddLog(@"***************************************************************", true);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                }
                else
                {
                    //有以前记录，需要看是否超差 并且 超差不通过 并且 超差不通过的提示选择不生成证书 时选择退出
                    if (needTestGeCe && !canGeCe && (MessageBox.Show("检测到有数据超差，是否继续生成证书？", "问题", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No))
                    {
                        AddException("有实验数据超差，暂时保留原记录，不做任何处理。", true);
                        if (_sr != null && _sr.ExcelWorkbook != null)
                        {
                            _sr.ExcelWorkbook.Saved = true;
                            _sr.TryClose();
                        }
                        if (_eu != null && _eu.ExcelWorkbook != null)
                        {
                            _eu.ExcelWorkbook.Saved = true;
                            _eu.TryClose();
                        }
                        DataUtility.DataUtility.RestoreFile(tempFo, backupKey);
                        problemFilesList.Add(fileText);
                        AddLog(@"***************************************************************", true);
                        exceptionNum = 0;
                        dataerrorNum = 0;
                    }
                    else
                    {
                        bool success = false;
                        //关闭归档后的Excel
                        if (_eu != null && _eu.ExcelWorkbook != null)
                        {
                            _eu.ExcelWorkbook.Saved = true;
                            _eu.TryClose();
                        }
                        //新文档因为将sr另存为合并后的文档，需要关闭后重新打开原文档
                        if (!needTestGeCe)
                        {
                            //关闭原来另存为成合并好的excel文件
                            if (_sr != null && _sr.ExcelWorkbook != null)
                            {
                                _sr.ExcelWorkbook.Saved = true;
                                _sr.TryClose();
                            }
                            //重新打开待合并的excel文件
                            _sr = new ExcelUtility(filePath, out checkClear);
                            if (!checkClear)
                            {
                                AddException(@"Excel文档无法打开", true);
                                if (_sr != null && _sr.ExcelWorkbook != null)
                                {
                                    _sr.ExcelWorkbook.Saved = true;
                                    _sr.TryClose();
                                }
                                AddLog(@"***************************************************************", true);
                                problemFilesList.Add(filePath);
                                exceptionNum = 0;
                                dataerrorNum = 0;
                                return;
                            }
                            //打开成功
                            _sr.ExcelApp.DisplayAlerts = false;
                            _sr.ExcelApp.AlertBeforeOverwriting = false;
                        }

                        //找到原记录里的数据页，记下序号
                        stateIndex = -1;
                        string temp = "";
                        foreach (MSExcel.Worksheet item in _sr.ExcelWorkbook.Sheets)
                        {
                            temp = _sr.GetText(_sr.ExcelWorkbook, item.Index, "L2", out checkClear);
                            if (certId == item.Name || certId == temp)
                            {
                                stateIndex = item.Index;
                            }
                        }
                        //找到序号的话，加入校核人的签名，删除其他sheet
                        if (stateIndex > 0 && stateIndex < _sr.ExcelWorkbook.Worksheets.Count)
                        {
                            _sr.WriteImage(_sr.ExcelWorkbook, stateIndex, 29, 9, person, 45, 28, out success);
                            for (int i = _sr.ExcelWorkbook.Worksheets.Count; i > 0; i--)
                            {
                                if (i != stateIndex)
                                {
                                    ((MSExcel.Worksheet)_sr.ExcelApp.ActiveWorkbook.Sheets[i]).Delete();
                                }
                            }
                            _sr.ExcelWorkbook.Save();
                            _sr.ExcelWorkbook.Saved = true;
                        }

                        //出证书
                        GenerateCert(_sr, stateIndex, pS.DataPattern, pS.CertTemplateFilePath, pS.CertFolder, pS.PDFDataFolder, pS.TempFolder, shouldFix, out success);
                        if (!success)
                        {
                            AddException("生成证书失败", true);
                            if (_sr != null && _sr.ExcelWorkbook != null)
                            {
                                _sr.ExcelWorkbook.Saved = true;
                                _sr.TryClose();
                            }
                            problemFilesList.Add(fileText);
                            AddLog(@"***************************************************************", true);
                            exceptionNum = 0;
                            dataerrorNum = 0;
                        }
                        else
                        {
                            if (_sr != null && _sr.ExcelWorkbook != null)
                            {
                                _sr.ExcelWorkbook.Saved = true;
                                _sr.TryClose();
                            }
                            File.Delete(filePath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddException(@"生成证书时遇到异常：" + ex.Message, true);
            }
            if (dataerrorNum > 0)
            {
                dataerrorNum = 0;
                AddLog(@"***************************************************************", true);
            }
        }

        #endregion
        #endregion

        #region HistoryFunction
        #region Standerdize

        public void StandardizeOne(string filePath, JobParameterStruct pS, ref List<string> problemFilesList)
        {
            int pattern = pS.DataPattern;
            string output = pS.OutputFolder;
            string templateName = pS.DataTemplateFilePath;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();
            Dictionary<int, string> cert = new Dictionary<int, string>();
            List<string> sheetsName = new List<string>();

            int stateIndex = -1;
            int noIdNumber = 0;
            int startDestiRowIndex = -1;

            bool Perfect;
            bool needFix = true;
            bool shouldFix = true;
            bool hasTemplate = false;
            bool canDelete = false;

            string year = "";
            string certId;
            string tempName;
            string text;
            string temp_strCompany = "", temp_strType = "", temp_strMacSerial = "", temp_strSensorSerial = "";
            string filePathNew = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + "_处理中" + Path.GetExtension(filePath));

            FileInfo temp_fi = null;

            File.Move(filePath, filePathNew);
            ExcelUtility _sr = new ExcelUtility(filePathNew, out checkClear);
            if (!checkClear)
            {
                AddException(@"Excel文档无法打开", true);
                if (_sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            _sr.ExcelApp.DisplayAlerts = false;
            _sr.ExcelApp.AlertBeforeOverwriting = false;
            MSExcel.Range rr = null;

            sheetsName = new List<string>();
            exSheets = new Dictionary<int, string>();

            //第一次循环：获取信息，并规范每页的标签
            foreach (MSExcel.Worksheet item in _sr.ExcelWorkbook.Sheets)
            {
                if (!hasTemplate && item.Name == @"标准模板")
                {
                    hasTemplate = true;
                }

                //规范sheet标签名为证书编号
                certId = _sr.GetText(_sr.ExcelWorkbook, item.Index, "L2", out checkClear).Trim();
                if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                {
                    //有规范的证书号
                    exSheets.Add(item.Index, certId);
                    stateIndex = item.Index;
                }
                else
                {
                    //无规范的证书号
                    rr = _sr.GetRange(_sr.ExcelWorkbook, item.Index, "A4", out checkClear);
                    if (!item.Name.Contains(@"标准模板") && rr.Text.ToString().Trim().StartsWith(@"送校单位"))
                    {
                        //有记录不包含规范的证书编号
                        AddException(@"该文档有实验数据不包含证书编号", true);
                        noIdNumber++;
                    }
                }
            }

            if (exSheets.Count == 0)
            {
                if (noIdNumber == 0)
                {
                    AddException(@"该文档可能是空文档", true);
                }
                if (_sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            else if (exSheets.Count + noIdNumber > 1)
            {
                AddException(@"该文档包含多个数据sheet，默认处理第一个", true);
            }

            TypeStandardize(_sr, stateIndex);

            certId = exSheets[stateIndex];
            year = certId.Substring(0, 4);

            tempName = year + "_" + _sr.GenerateFileName(_sr.ExcelWorkbook, stateIndex, out temp_strCompany, out temp_strType, out temp_strMacSerial, out temp_strSensorSerial, out Perfect);
            if (temp_strCompany == "") AddException(@"送校单位信息未提取到", true);
            if (temp_strType == "") AddException(@"仪器型号信息未提取到", true);
            if ((temp_strSensorSerial.Length + temp_strMacSerial.Length) == 0) AddException(@"主机编号和探头编号信息未提取到", true);
            if (temp_strMacSerial == "" || temp_strMacSerial == "/") AddException(@"主机编号信息未提取到", true);

            GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
            switch (pattern)
            {
                case 0:
                    //TODO : 剂量 更改路径
                    startDestiRowIndex = 18;
                    //templateName = DataUtility.DataUtility.PathClassifiedCombineName(Application.StartupPath + @"\试验记录模板", pattern, @"剂量证书实验记录模板.xlsx");
                    break;
                case 1:
                    //TODO : CT 更改路径
                    startDestiRowIndex = 18;
                    //templateName = DataUtility.DataUtility.PathClassifiedCombineName(Application.StartupPath + @"\试验记录模板", pattern, @"ct证书实验记录模板.xlsx");
                    break;
                case 2:
                    //TODO : KV 更改路径
                    startDestiRowIndex = 17;
                    //templateName = DataUtility.DataUtility.PathClassifiedCombineName(Application.StartupPath + @"\试验记录模板", pattern, @"KV校准证书实验记录模板.xlsx");
                    break;
                default:
                    AddException(@"数据类型无效", true);
                    if (_sr.ExcelWorkbook != null)
                    {
                        _sr.ExcelWorkbook.Saved = true;
                        _sr.TryClose();
                    }
                    AddLog(@"***************************************************************", true);
                    problemFilesList.Add(filePath);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                    return;
            }

            if (File.Exists(templateName))
            {
                tempName = DataUtility.DataUtility.PathRightFileName(output, tempName, ".xlsx", "_new");

                File.Copy(templateName, tempName, true);
                if (!File.Exists(tempName))
                {
                    AddException(@"拷贝模板失败", true);
                    if (_sr.ExcelWorkbook != null)
                    {
                        _sr.ExcelWorkbook.Saved = true;
                        _sr.TryClose();
                    }
                    AddLog(@"***************************************************************", true);
                    problemFilesList.Add(filePath);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                    return;
                }
                else
                {
                    temp_fi = new FileInfo(tempName);
                }
            }
            else
            {
                AddException(@"模板不存在：" + templateName, true);
                if (_sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }

            exSheets.Clear();

            ExcelUtility _eu = new ExcelUtility(temp_fi.FullName, out checkClear);
            if (!checkClear)
            {
                AddException(@"Excel文档无法打开", true);
                if (_sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                if (_eu.ExcelWorkbook != null)
                {
                    _eu.ExcelWorkbook.Saved = true;
                    _eu.TryClose();
                }
                File.Delete(temp_fi.FullName);

                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            _eu.ExcelApp.DisplayAlerts = false;
            _eu.ExcelApp.AlertBeforeOverwriting = false;
            Object format = MSExcel.XlFileFormat.xlWorkbookDefault;

            try
            {
                int newSheetIndex = -1;
                CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out newSheetIndex, out checkClear);
                if (newSheetIndex > 0 && checkClear)
                {
                    bool testflag;
                    text = _eu.GetText(_eu.ExcelWorkbook, newSheetIndex, 2, 11, out testflag);
                    if (!text.ToLower().EndsWith("dyjl"))
                    {
                        AddException("证书编号不以DYjl开始", true);
                    }

                    if (certId.StartsWith("200"))
                    {
                        _eu.WriteValue(_eu.ExcelWorkbook, newSheetIndex, 2, 11, "证书编号：DYjx", "", out testflag);
                    }
                }
                else
                {
                    AddException(@"无法找到新合并入数据的位置", true);
                }
                _eu.ExcelWorkbook.Save();
            }
            catch (Exception ex)
            {
                AddException(@"标准化时遇到异常：" + ex.Message, true);
            }
            finally
            {
                if (_sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                if (_eu.ExcelWorkbook != null)
                {
                    _eu.ExcelWorkbook.Saved = true;
                    _eu.TryClose();
                }
                if (exceptionNum > 0)
                {
                    File.Delete(temp_fi.FullName);

                    AddLog(@"***************************************************************", true);
                    problemFilesList.Add(filePath);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                }
                else
                {
                    canDelete = true;
                }
            }

            if (canDelete)
            {
                try
                {
                    File.Delete(filePathNew);
                }
                catch (Exception ex)
                {
                    AddDataError("已处理完成，但原文件无法正常删除！遇到异常：" + ex.Message, true);
                }
            }
        }

        #endregion

        #region ArchievingMerge
        /// <summary>
        /// 存档合并模式
        /// </summary>
        /// <param name="filePath">需要处理的文件名</param>
        /// <param name="pS">处理过程需要的参数集</param>
        /// <param name="problemFilesList">出问题的文件列表</param>
        public void ArchievingMergeOne(string filePath, JobParameterStruct pS, ref List<string> problemFilesList)
        {
            int pattern = pS.DataPattern;
            string output = pS.OutputFolder;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();

            bool Perfect;
            bool isArchieved = false;
            bool needFix = true;
            bool shouldFix = true;

            int stateIndex = -1;
            int noIdNumber = 0;
            int startDestiRowIndex = -1;

            string tempName;
            string certId;
            string newName = "";
            string fileText = "";
            string renameStr;
            string strCompany = "", strType = "", strMacSerial = "", strSensorSerial = "";

            FileInfo temp_fi = null;
            FileInfo fi = new FileInfo(filePath);

            Object format = MSExcel.XlFileFormat.xlWorkbookDefault;

            existFile = (new DirectoryInfo(output)).GetFiles(@"*.xls*", SearchOption.AllDirectories);

            ExcelUtility _sr = new ExcelUtility(filePath, out checkClear);
            ExcelUtility _eu = null;
            MSExcel.Range rr = null;
            if (!checkClear)
            {
                AddException(@"Excel文档无法打开", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            _sr.ExcelApp.DisplayAlerts = false;
            _sr.ExcelApp.AlertBeforeOverwriting = false;

            //第一次循环：获取信息，并规范每页的标签
            foreach (MSExcel.Worksheet item in _sr.ExcelWorkbook.Sheets)
            {
                //规范sheet标签名为证书编号
                certId = _sr.GetText(_sr.ExcelWorkbook, item.Index, "L2", out checkClear).Trim();
                if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                {
                    //有规范的证书号
                    exSheets.Add(item.Index, certId);
                    stateIndex = item.Index;
                }
                else
                {
                    //无规范的证书号
                    rr = _sr.GetRange(_sr.ExcelWorkbook, item.Index, "A4", out checkClear);
                    if (!item.Name.Contains(@"标准模板") && rr.Text.ToString().Trim().StartsWith(@"送校单位"))
                    {
                        //有记录不包含规范的证书编号
                        AddException(@"该文档有实验数据不包含证书编号", true);
                        noIdNumber++;
                    }
                }
                renameStr = _sr.GetText(_sr.ExcelWorkbook, item.Index, 5, 7, out checkClear).Trim().ToLower();
                if (renameStr.StartsWith(@"编号："))
                {
                    _sr.WriteValue(_sr.ExcelWorkbook, item.Index, 5, 7, renameStr.Replace(@"编号：", @"主机编号："), out checkClear);
                }

                renameStr = _sr.GetText(_sr.ExcelWorkbook, item.Index, 5, 11, out checkClear).Trim().ToLower();
                if (renameStr.StartsWith(@"电离室号："))
                {
                    _sr.WriteValue(_sr.ExcelWorkbook, item.Index, 5, 11, renameStr.Replace(@"电离室号：", @"探测器编号："), out checkClear);
                }
                else if (renameStr.StartsWith(@"探测器号："))
                {
                    _sr.WriteValue(_sr.ExcelWorkbook, item.Index, 5, 11, renameStr.Replace(@"探测器号：", @"探测器编号："), out checkClear);
                }
            }
            _sr.ExcelWorkbook.Save();
            _sr.ExcelWorkbook.Saved = true;

            if (exSheets.Count == 0)
            {
                if (noIdNumber == 0)
                {
                    AddException(@"该文档可能是空文档", true);
                }
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            else if (exSheets.Count + noIdNumber > 1)
            {
                AddException(@"该文档包含多个数据sheet，默认处理第一个", true);
            }

            certId = exSheets[stateIndex];
            tempName = _sr.GenerateFileName(_sr.ExcelWorkbook, stateIndex, out strCompany, out strType, out strMacSerial, out strSensorSerial, out Perfect);
            checkClear = false;
            if (strCompany == "")
            {
                AddException(@"送校单位信息未提取到", true);
                checkClear = true;
            }
            if (strType == "")
            {
                AddException(@"仪器型号信息未提取到", true);
                checkClear = true;
            }
            if (strMacSerial == "" || strMacSerial == "主机_/")
            {
                AddException(@"主机编号信息未提取到", true);
                checkClear = true;
            }
            if (checkClear)
            {
                AddLog(@"***************************************************************", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
            if (pattern < 2)
            {
                startDestiRowIndex = 18;
            }
            else if (pattern == 2)
            {
                startDestiRowIndex = 17;
            }
            else
            {
                AddException(@"数据类型无效", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }

            try
            {
                //寻找目标统计文件
                temp_fi = SearchForFile(filePath, strType, strMacSerial, strSensorSerial, existFile, out isArchieved);
                if (temp_fi == null)
                {
                    if (isArchieved)
                    {
                        //不存在，复制当前记录过去
                        //1.合成新文件名，按新名另存，定义为temp_fi
                        newName = DataUtility.DataUtility.PathCombineFolderFileExtension(output, DataUtility.DataUtility.FileNameCleanName(tempName), fi.Extension);
                        fileText = newName;
                        _sr.ExcelWorkbook.SaveAs(newName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        temp_fi = new FileInfo(newName);
                        //2.统计新文件
                        Statistic(_sr, pattern, Perfect, strCompany, strType, tempName, certId);
                        //3.对新文件进行校验
                        _sr.Verification(_sr.ExcelWorkbook, true, pattern, out checkClear);
                    }
                    else
                    {
                        //不存在，但搜索到有可疑项，暂不处理
                        fileText = filePath;
                    }
                }
                else if (File.Exists(temp_fi.FullName))
                {
                    fileText = temp_fi.FullName;
                    _eu = new ExcelUtility(temp_fi.FullName, out checkClear);
                    if (!checkClear)
                    {
                        AddException(@"Excel文档无法打开", true);
                        _sr.ExcelWorkbook.Saved = true;
                        _eu.ExcelWorkbook.Saved = true;
                        _sr.TryClose();
                        _eu.TryClose();
                        AddLog(@"***************************************************************", true);
                        problemFilesList.Add(filePath);
                        exceptionNum = 0;
                        dataerrorNum = 0;
                        return;
                    }
                    _eu.ExcelApp.DisplayAlerts = false;
                    _eu.ExcelApp.AlertBeforeOverwriting = false;

                    CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out checkClear);
                    if (checkClear)
                    {
                        _eu.ExcelWorkbook.Save();
                        _eu.ExcelWorkbook.Saved = true;
                        Statistic(_eu, pattern, Perfect, strCompany, strType, tempName, certId);
                        _eu.ExcelWorkbook.Save();
                        _eu.ExcelWorkbook.Saved = true;
                        //TODO: 存档合并处校验是否可选（存在，合并入原记录）
                        _eu.Verification(_eu.ExcelWorkbook, false, pattern, out checkClear);
                    }
                }
                else
                {
                    AddException(@"文件不存在：" + temp_fi.FullName, true);
                    fileText = filePath;
                }
            }
            catch (Exception ex)
            {
                AddException(@"存档合并时遇到异常：" + ex.Message, true);
            }
            finally
            {
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                if (_eu != null && _eu.ExcelWorkbook != null)
                {
                    _eu.ExcelWorkbook.Saved = true;
                    _eu.TryClose();
                }
                if (exceptionNum > 0)
                {
                    problemFilesList.Add(fileText);
                    AddLog(@"***************************************************************", true);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                }
                else
                {
                    if (dataerrorNum > 0)
                    {
                        AddLog(@"该文件只有非致命错误，已经成功处理", true);
                        AddLog(@"***************************************************************", true);
                        exceptionNum = 0;
                        dataerrorNum = 0;
                    }
                    File.Delete(filePath);
                }

            }
        }

        #endregion

        #region Verificate
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="pattern"></param>
        /// <param name="text1"></param>
        /// <param name="text7"></param>
        /// <param name="text3"></param>
        /// <param name="problemFilesList"></param>
        public void VerificateOne(string filePath, JobParameterStruct pS, ref List<string> problemFilesList)
        {
            int pattern = pS.DataPattern;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();
            Dictionary<int, string> cert = new Dictionary<int, string>();
            List<string> sheetsName = new List<string>();
            string fileName = "";
            bool pass = false;

            ExcelUtility _sr = new ExcelUtility(filePath, out checkClear);
            if (!checkClear)
            {
                AddException(@"Excel文档无法打开", true);
                _sr.ExcelWorkbook.Saved = true;
                _sr.TryClose();
                AddLog(@"***************************************************************", true);
                problemFilesList.Add(filePath);
                exceptionNum = 0;
                dataerrorNum = 0;
                return;
            }
            _sr.ExcelApp.DisplayAlerts = false;
            _sr.ExcelApp.AlertBeforeOverwriting = false;

            try
            {
                //TODO: 单独校验处校验是否可选
                fileName = _sr.Verification(_sr.ExcelWorkbook, true, pattern, out pass);
            }
            catch (Exception ex)
            {
                AddException(@"校验时遇到异常：" + ex.Message, true);
            }
            finally
            {
                if (pass)
                {
                    FileInfo fi = new FileInfo(filePath);
                    fileName = DataUtility.DataUtility.PathRightFileName(fi.DirectoryName, fileName, fi.Extension, "_new");
                    if (filePath != fileName)
                    {
                        _sr.ExcelWorkbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Thread.Sleep(1000);
                        File.Delete(filePath);
                    }
                    else
                    {
                        _sr.ExcelWorkbook.Save();
                    }
                }
                _sr.ExcelWorkbook.Saved = true;
                _sr.TryClose();

                if (exceptionNum > 0)
                {
                    AddLog(@"***************************************************************", true);
                    problemFilesList.Add(filePath);
                    exceptionNum = 0;
                    dataerrorNum = 0;
                }
            }
        }

        #endregion

        #endregion
    }
}


//public bool Statistic(ExcelUtility eu, int pattern, bool Perfect, string strCompany, string strType, string tempName)
//{
//    int stateIndex = -1;
//    int lineIndex = 0;
//    int limitPosition = 14;

//    bool needInitialState = true;
//    bool checkClear;
//    bool firstTime = true;
//    bool canPass = true, canPass1 = false;

//    string insNumber = "";

//    Dictionary<int, string> exSheets = null; 
//    Dictionary<int, string> cert = new Dictionary<int, string>();
//    Dictionary<Criterion, int> existCriList = new Dictionary<Criterion,int>();

//    //1.初始化：规范统计sheet结构，写入固定内容
//    stateIndex = eu.InitialStatisticSheet(eu.ExcelWorkbook, out exSheets, out needInitialState, out existCriList, out limitPosition, out checkClear);
//    if (!checkClear) { return false; }
//    if (needInitialState)
//    {
//        eu.WriteStateTitle(eu.ExcelWorkbook, pattern, stateIndex, existCriList, limitPosition);
//        //写入送校单位
//        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 2, 2, strCompany, out checkClear);
//        //写入仪器名称
//        eu.WriteValue(eu.ExcelWorkbook, stateIndex, 3, 2, strType, out checkClear);
//    }

//    switch (pattern)
//    {
//        case 0:
//            //1.搬运数据：先根据年份排序，然后复制数据
//            int[] indexDose = new int[exSheets.Count];
//            foreach (int item in exSheets.Keys)
//            {
//                indexDose[lineIndex] = item;
//                lineIndex++;
//            }
//            DataUtility.DataUtility.QuickSort(indexDose, exSheets, 0, indexDose.Length - 1);

//            lineIndex = 8;
//            for (int i = 0; i < exSheets.Count; i++ )
//            {
//                cert.Add(lineIndex, exSheets[indexDose[i]]);
//                eu.CopyDoseOneYearData(eu.ExcelWorkbook, indexDose[i], stateIndex, lineIndex, exSheets[indexDose[i]], limitPosition, out insNumber);

//                if (firstTime && Perfect)
//                {
//                    eu.DesiredName = tempName;

//                    firstTime = false;
//                }
//                lineIndex++;
//            }                    
//            //2.分析数据

//            if (exSheets.Count > 1)
//            {
//                eu.WriteValue(eu.ExcelWorkbook, stateIndex, lineIndex, 1, @"重复性", out checkClear);
//                eu.ExcelWorkbook.Save();
//                eu.ExcelWorkbook.Saved = true;
//                //计算平均值和年稳定性
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR4（60kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 6, cert, "RQR5（70kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR6（80kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 10, cert, "RQR8（100kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR9（120kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//            }
//            break;
//        case 1:
//            //TODO: CT统计代码
//            //1.搬运数据：先根据年份排序，然后复制数据
//            int[] indexCT = new int[exSheets.Count];
//            foreach (int item in exSheets.Keys)
//            {
//                indexCT[lineIndex] = item;
//                lineIndex++;
//            }
//            DataUtility.DataUtility.QuickSort(indexCT, exSheets, 0, indexCT.Length - 1);

//            lineIndex = 8;
//            for (int i = 0; i < exSheets.Count; i++)
//            {
//                cert.Add(lineIndex, exSheets[indexCT[i]]);
//                //TODO:复制数据格式
//                eu.CopyCTOneYearData(eu.ExcelWorkbook, indexCT[i], stateIndex, lineIndex, exSheets[indexCT[i]], limitPosition, existCriList, out insNumber);

//                if (firstTime && Perfect)
//                {
//                    eu.DesiredName = tempName;

//                    firstTime = false;
//                }
//                lineIndex++;
//            }
//            //2.分析数据

//            if (exSheets.Count > 1)
//            {
//                eu.WriteValue(eu.ExcelWorkbook, stateIndex, lineIndex, 1, @"重复性", out checkClear);
//                eu.ExcelWorkbook.Save();
//                eu.ExcelWorkbook.Saved = true;
//                //计算平均值和年稳定性
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 4, cert, "RQR4（60kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 6, cert, "RQR5（70kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 8, cert, "RQR6（80kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 10, cert, "RQR8（100kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//                canPass1 = eu.StatisticsOneColumn(eu.ExcelWorkbook, stateIndex, lineIndex, 12, cert, "RQR9（120kV）", exSheets.Count > 1, true);
//                canPass = canPass1 && canPass;
//            }
//            break;
//        case 2:
//            //TODO：KV统计代码
//            break;
//    }
//    eu.ExcelWorkbook.Save();
//    return canPass;
//}
