using System;
using System.Collections.Generic;
using System.IO;

using MSExcel = Microsoft.Office.Interop.Excel;

using Statistics.Configuration;
using Statistics.Log;
using Statistics.Office.Excel;
using Statistics.SuperDog;
using Util;

namespace Statistics.Job
{
    public class GeneratingCertificateJob : Job
    {
        public GeneratingCertificateJob(string filePath, JobParameterStruct jobParam, Person person)
            : base(filePath, jobParam, person)
        {

        }

        public override void DoTheJob()
        {
            int pattern = jobParam.DataPattern;
            string output = jobParam.AutoOutputFolder;
            string ext = jobParam.AutoExtension;
            bool createNew = jobParam.CreateNew;
            string tempFo = jobParam.TempFolder;

            FileInfo[] existFile = null;

            Dictionary<int, string> exSheets = new Dictionary<int, string>();

            bool Perfect;
            bool success;
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

            ExcelUtility _sr = new ExcelUtility(filePath, out success);
            ExcelUtility _eu = null;
            MSExcel.Range rr = null;
            if (!success)
            {
                LogHelper.AddException(@"Excel文档无法打开", true);
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                LogHelper.AddProblemFilesAndReset(filePath);
                return;
            }
            _sr.ExcelApp.DisplayAlerts = false;
            _sr.ExcelApp.AlertBeforeOverwriting = false;

            //第一次循环：获取信息，并规范每页的标签
            foreach (MSExcel.Worksheet item in _sr.ExcelWorkbook.Sheets)
            {
                //规范sheet标签名为证书编号
                certId = _sr.GetText(_sr.ExcelWorkbook, item.Index, new ExcelPosition("L2"));
                if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                {
                    //有规范的证书号
                    exSheets.Add(item.Index, certId);
                    stateIndex = item.Index;
                }
                else
                {
                    //无规范的证书号
                    rr = _sr.GetRange(_sr.ExcelWorkbook, item.Index, new ExcelPosition("A4"));
                    if (!item.Name.Contains(@"标准模板") && rr.Text.ToString().Trim().StartsWith(@"送校单位"))
                    {
                        //有记录不包含规范的证书编号
                        LogHelper.AddException(@"该文档有实验数据不包含证书编号", true);
                        noIdNumber++;
                    }
                }
            }

            if (exSheets.Count == 0)
            {
                if (noIdNumber == 0)
                {
                    LogHelper.AddException(@"该文档可能是空文档", true);
                }
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                LogHelper.AddProblemFilesAndReset(filePath);
                return;
            }
            else if (exSheets.Count + noIdNumber > 1)
            {
                LogHelper.AddException(@"该文档包含多个数据sheet，默认处理第一个", true);
            }

            certId = exSheets[stateIndex];
            tempName = _sr.GenerateFileName(_sr.ExcelWorkbook, stateIndex, out strCompany, out strType, out strMacSerial, out strSensorSerial, out Perfect);
            success = false;
            if (strCompany == "")
            {
                LogHelper.AddException(@"送校单位信息未提取到", true);
                success = true;
            }
            if (strType == "")
            {
                LogHelper.AddException(@"仪器型号信息未提取到", true);
                success = true;
            }
            if (strMacSerial == "")
            {
                LogHelper.AddException(@"主机编号信息未提取到", true);
                success = true;
            }
            if (success)
            {
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                LogHelper.AddProblemFilesAndReset(filePath);
                return;
            }
            JobMethods.GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
            //目标行数
            startDestiRowIndex = 18;

            try
            {
                //寻找目标统计文件
                existFile = JobMethods.GetFilesFromType(output, _sr.GetText(_sr.ExcelWorkbook, stateIndex, new ExcelPosition("F5")), ext, out success);
                if (!success)
                {
                    if (_sr != null && _sr.ExcelWorkbook != null)
                    {
                        _sr.ExcelWorkbook.Saved = true;
                        _sr.TryClose();
                    }
                    LogHelper.AddProblemFilesAndReset(filePath);
                    return;
                }

                temp_fi = JobMethods.SearchForFile(filePath, strType, strMacSerial, strSensorSerial, existFile, out hasArchieved);
                if (temp_fi == null)
                {
                    if (createNew)
                    {
                        //对新记录建档
                        needTestGeCe = false;
                        if (hasArchieved)
                        {
                            //不存在，复制当前记录过去
                            newName = DataUtility.DataUtility.PathRightFileName(Util.PathExt.PathCombine(output, strType), tempName, fi.Extension, "_new");
                            _sr.ExcelWorkbook.SaveAs(newName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            temp_fi = new FileInfo(newName);
                            JobMethods.Statistic(_sr, pattern, Perfect, strCompany, strType, tempName, certId);
                            //TODO: 生成证书处校验是否可选（不存在，复制当前记录过去）
                            _sr.Verification(_sr.ExcelWorkbook, true, pattern, out success);
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
                        LogHelper.AddException("没有在历史数据记录中发现匹配项，暂不处理。", true);
                        if (_sr != null && _sr.ExcelWorkbook != null)
                        {
                            _sr.ExcelWorkbook.Saved = true;
                            _sr.TryClose();
                        }
                        LogHelper.AddProblemFilesAndReset(filePath);
                        return;
                    }
                }
                else if (File.Exists(temp_fi.FullName))
                {
                    DataUtility.DataUtility.BackupFile(tempFo, temp_fi.FullName, out backupKey);
                    _eu = new ExcelUtility(temp_fi.FullName, out success);
                    if (!success)
                    {
                        LogHelper.AddException(@"Excel文档无法打开", true);
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
                        LogHelper.AddProblemFilesAndReset(filePath);
                        return;
                    }
                    _eu.ExcelApp.DisplayAlerts = false;
                    _eu.ExcelApp.AlertBeforeOverwriting = false;

                    JobMethods.CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out success);
                    if (!success)
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
                        LogHelper.AddProblemFilesAndReset(fileText);
                        return;
                    }
                    _eu.ExcelWorkbook.Save();
                    _eu.ExcelWorkbook.Saved = true;
                    canGeCe = JobMethods.Statistic(_eu, pattern, Perfect, strCompany, strType, tempName, certId);
                    _eu.ExcelWorkbook.Save();
                    _eu.ExcelWorkbook.Saved = true;
                    //TODO: 生成证书处校验是否可选（存在，合并入原记录）
                    _eu.Verification(_eu.ExcelWorkbook, false, pattern, out success);
                    fileText = temp_fi.FullName;
                }
                else
                {
                    LogHelper.AddException(@"文件不存在：" + temp_fi.FullName, true);
                    fileText = filePath;
                }
            }
            catch (Exception ex)
            {
                LogHelper.AddException(@"生成证书时合并一步遇到异常：" + ex.Message, true);
            }

            try
            {
                //有重大失误，关闭两个excel，报错退出。没有失误继续运行
                if (LogHelper.HasException)
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
                    LogHelper.AddProblemFilesAndReset(fileText);
                }
                else
                {
                    //有以前记录，需要看是否超差 并且 超差不通过 并且 超差不通过的提示选择不生成证书 时选择退出
                    if (needTestGeCe && !canGeCe && !FormOperator.MessageBox_Show_YesNo("检测到有数据超差，是否继续生成证书？", "问题"))
                    {
                        LogHelper.AddException("有实验数据超差，暂时保留原记录，不做任何处理。", true);
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
                        LogHelper.AddProblemFilesAndReset(fileText);
                    }
                    else
                    {
                        success = false;
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
                            _sr = new ExcelUtility(filePath, out success);
                            if (!success)
                            {
                                LogHelper.AddException(@"Excel文档无法打开", true);
                                if (_sr != null && _sr.ExcelWorkbook != null)
                                {
                                    _sr.ExcelWorkbook.Saved = true;
                                    _sr.TryClose();
                                }
                                LogHelper.AddProblemFilesAndReset(filePath);
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
                            temp = _sr.GetText(_sr.ExcelWorkbook, item.Index, new ExcelPosition("L2"));
                            if (certId == item.Name || certId == temp)
                            {
                                stateIndex = item.Index;
                            }
                        }
                        //找到序号的话，加入校核人的签名，删除其他sheet
                        if (stateIndex > 0 && stateIndex < _sr.ExcelWorkbook.Worksheets.Count)
                        {
                            _sr.WriteImage(_sr.ExcelWorkbook, stateIndex, new ExcelPosition(30, 9), Util.PathExt.PathCombine(ProgramConfiguration.ProgramFolder, person.Path), 45, 28);
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
                        GenerateCert(_sr, stateIndex, jobParam.DataPattern, jobParam.CertTemplateFilePath, jobParam.CertFolder, jobParam.PDFDataFolder, jobParam.TempFolder, shouldFix, out success);
                        if (!success)
                        {
                            LogHelper.AddException("生成证书失败", true);
                            if (_sr != null && _sr.ExcelWorkbook != null)
                            {
                                _sr.ExcelWorkbook.Saved = true;
                                _sr.TryClose();
                            }
                            LogHelper.AddProblemFilesAndReset(fileText);
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
                LogHelper.AddException(@"生成证书时遇到异常：" + ex.Message, true);
            }
            if (LogHelper.HasDataError)
            {
                LogHelper.ResetError();
                LogHelper.AddLog(@"***************************************************************", true);
            }
        }

        public override JobType JobType
        {
            get { return JobType.GeneratingCertificate; }
        }

        #region Private Method
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
                    Log.LogHelper.AddException("Word文档打开失败", true);
                    return;
                }
                string stemp1 = excel.GetText(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("L2"));
                object otemp1;
                string wdName = "DYjl" + stemp1 + Path.GetExtension(wordPath);
                string pdfName = "DYjl" + stemp1 + "_" + excel.GetText(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("B4")) + ".pdf";

                switch (pattern)
                {
                    case 0:
                        //检定
                        //检定依据
                        wu.WriteValue(wu.WordDocument, "m_JDYJ", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("B8")));
                        //有效日期
                        string[] dates = excel.GetText(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("K29")).Split(new string[] { "年", "月", "日", " " }, StringSplitOptions.RemoveEmptyEntries);
                        wu.WriteValue(wu.WordDocument, "m_YXRQ1", dates[0].Substring(2));
                        wu.WriteValue(wu.WordDocument, "m_YXRQ2", dates[1]);
                        wu.WriteValue(wu.WordDocument, "m_YXRQ3", dates[2]);
                        //检定日期
                        dates = excel.GetText(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("K30")).Split(new string[] { "年", "月", "日", " " }, StringSplitOptions.RemoveEmptyEntries);
                        wu.WriteValue(wu.WordDocument, "m_JDRQ1", dates[0].Substring(2));
                        wu.WriteValue(wu.WordDocument, "m_JDRQ2", dates[1]);
                        wu.WriteValue(wu.WordDocument, "m_JDRQ3", dates[2]);
                        //备注
                        wu.WriteValue(wu.WordDocument, "m_BZ", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("B27")));
                        break;
                    case 1:
                        //校准
                        //客户地址
                        wu.WriteValue(wu.WordDocument, "m_KHDZ", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("F4")));
                        //校准依据
                        wu.WriteValue(wu.WordDocument, "m_JZYJ", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("B8")));
                        //校准日期
                        wu.WriteValue(wu.WordDocument, "m_JZRQ", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("K30")));
                        //扩展不确定度
                        wu.WriteValue(wu.WordDocument, "m_KZBQDD", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("K26")));
                        break;
                    default:
                        Log.LogHelper.AddException("生成证书时指定了不存在的检定类型", true);
                        break;
                }

                /// <summary>
                /// 类型1：普通复制
                /// </summary>
                wu.WriteValue(wu.WordDocument, "m_SJDW", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("B4")));
                wu.WriteValue(wu.WordDocument, "m_QJMC", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("B5")));
                wu.WriteValue(wu.WordDocument, "m_XHGG", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("F5")));
                wu.WriteValue(wu.WordDocument, "m_SCCS", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("J5")));
                wu.WriteValue(wu.WordDocument, "m_LC", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("D12")));
                wu.WriteValue(wu.WordDocument, "m_QY", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("J8")));
                wu.WriteValue(wu.WordDocument, "m_WD", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("K7")));
                wu.WriteValue(wu.WordDocument, "m_ZSBH1", stemp1);
                wu.WriteValue(wu.WordDocument, "m_ZSBH2", stemp1);
                wu.WriteValue(wu.WordDocument, "m_ZSBH3", stemp1);

                /// <summary>
                /// 类型2：百分比换算后复制
                /// </summary>
                otemp1 = (object)excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("M7")).Value;
                if (otemp1 == null)
                {
                    stemp1 = "/";
                }
                else
                {
                    stemp1 = string.Format("{0:F1}", float.Parse(otemp1.ToString()) * 100);
                }
                wu.WriteValue(wu.WordDocument, "m_SD", stemp1);
                /// <summary>
                /// 类型3：仪器编号两段合并后复制
                /// </summary>
                otemp1 = (object)excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("H5")).Value;
                if (otemp1 == null)
                {
                    stemp1 = "";
                }
                else
                {
                    stemp1 = otemp1.ToString();
                }
                otemp1 = (object)excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("L5")).Value;
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
                wu.WriteValue(wu.WordDocument, "m_CCBH", stemp1);
                /// <summary>
                /// 类型4：数据部分
                /// </summary>
                wu.WriteDataValue(wu.WordDocument, "m_DATA1", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("D24")), "{0:F3}"); //小数点后三位
                wu.WriteDataValue(wu.WordDocument, "m_DATA2", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("F24")), "{0:F3}"); //小数点后三位
                wu.WriteDataValue(wu.WordDocument, "m_DATA3", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("H24")), "{0:F3}"); //小数点后三位
                wu.WriteDataValue(wu.WordDocument, "m_DATA4", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("J24")), "{0:F3}"); //小数点后三位
                wu.WriteDataValue(wu.WordDocument, "m_DATA5", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("L24")), "{0:F3}"); //小数点后三位
                wu.WriteDataValue(wu.WordDocument, "m_DATA6", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("D54")), "{0:F3}"); //小数点后三位
                wu.WriteDataValue(wu.WordDocument, "m_DATA7", excel.GetRange(excel.ExcelWorkbook, sourceIndex, new ExcelPosition("F54")), "{0:F3}"); //小数点后三位


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
        #endregion
    }
}
