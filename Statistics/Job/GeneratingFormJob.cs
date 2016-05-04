using System;
using System.Collections.Generic;
using System.IO;

using MSExcel = Microsoft.Office.Interop.Excel;

using Statistics.Configuration;
using Statistics.Log;
using Statistics.Office.Excel;
using Statistics.ProjectModel;
using Statistics.SuperDog;
using Util;

namespace Statistics.Job
{
    public class GeneratingFormJob : Job
    {
        public GeneratingFormJob(string filePath, JobParameterStruct jobParam, Person person)
            : base(filePath, jobParam, person)
        {

        }

        public override void DoTheJob()
        {
            string output = jobParam.AutoOutputFolder;
            FixType fix = jobParam.FixType;
            bool needFix = jobParam.AutoFixType;
            string templateName = jobParam.DataTemplateFilePath;
            string macType = jobParam.MacType;

            bool success = true;

            int templateIndex = -1;

            MSExcel.Worksheet ws1 = null;

            WordUtility _wu = new WordUtility(filePath, out success);
            if (!success)
            {
                LogHelper.AddException("Word文档打开失败", true);
                return;
            }

            string tempZhsh = _wu.GetText(_wu.WordDocument, 3);//L2:证书编号
            string tempName = _wu.GetText(_wu.WordDocument, 7);//B4:送校单位
            string tempQiju = _wu.GetText(_wu.WordDocument, 11);//B5:仪器名称
            string tempSerial = _wu.GetText(_wu.WordDocument, 15).Trim();//F5:仪器型号

            _wu.TryClose();

            if (tempSerial != "" && tempSerial != macType)
            {
                LogHelper.AddDataError("证书中包含的仪器型号与指定的仪器型号不符" + Environment.NewLine + "证书仪器型号: " + tempSerial + Environment.NewLine + "指定仪器型号: " + macType, true);
            }

            string str = tempZhsh.Substring(8);
            string strSavename = PathExt.PathCombine(output, tempName + "_" + macType + "_" + str + ".xlsx");

            if (File.Exists(strSavename))
            {
                if (FormOperator.MessageBox_Show_YesNo(@"文件已存在，是否覆盖？" + Environment.NewLine + strSavename, "提示"))
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

            ExcelUtility _sr = new ExcelUtility(strSavename, out success);
            if (!success)
            {
                if (_sr != null && _sr.ExcelWorkbook != null)
                {
                    _sr.ExcelWorkbook.Saved = true;
                    _sr.TryClose();
                }
                LogHelper.AddException(@"Excel文档无法打开", true);
                LogHelper.AddProblemFilesAndReset(filePath);
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
                        LogHelper.AddException(@"发现多余的标准模板", true);
                    }
                }
                LogHelper.State.Push("找到标准模板：" + templateIndex.ToString());
                if (templateIndex > -1)
                {
                    ws1 = (MSExcel.Worksheet)_sr.ExcelWorkbook.Sheets[templateIndex];
                    ws1.Copy(ws1, Type.Missing);
                    ws1 = (MSExcel.Worksheet)_sr.ExcelWorkbook.Sheets[templateIndex];
                    if (!ws1.Name.Contains(@"标准模板"))
                    {
                        LogHelper.AddException(@"标准模板复制出错", true);
                        success = false;
                        return;
                    }
                    else
                    {
                        ws1.Name = str;
                    }
                    LogHelper.State.Push("复制标准模板完成");
                }
                else
                {
                    LogHelper.AddException(@"找不到模板excel中的标准模板页", true);
                }

                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 4, 2, tempName, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 5, 6, macType, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 5, 2, tempQiju, out success);
                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 2, 12, str, out success);

                LogHelper.State.Push("写入4个信息");

                _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 8, 13, EnumExt.GetDescriptionFromEnumValue<FixType>(fix), out success);

                LogHelper.State.Push("写入修正信息");

                if (!needFix)
                {
                    //电离室->半导体
                    MSExcel.Range rr = _sr.GetRange(_sr.ExcelWorkbook, ws1.Index, new ExcelPosition("L8"));
                    rr.FormulaLocal = "";
                    rr.Formula = "";
                    rr.FormulaArray = "";

                    LogHelper.State.Push("不修正时清空公式");

                    _sr.WriteValue(_sr.ExcelWorkbook, ws1.Index, 8, 12, "1.000000", "@", out success);

                    LogHelper.State.Push("写入1");

                }

                LogHelper.State.Push("准备写入记录者图片");

                //写入记录者
                _sr.WriteImage(_sr.ExcelWorkbook, ws1.Index, 29, 7, PathExt.PathCombine(ProgramConfiguration.ProgramFolder, person.Path), 45, 28, out success);
                LogHelper.State.Push("写完记录者图片");
                _sr.ExcelWorkbook.Save();
                LogHelper.State.Push("保存完毕");
            }
            catch (Exception ex)
            {
                Log.LogHelper.AddException("生成证书时遇到异常：" + ex.Message, true);
                Log.LogHelper.AddLog("执行位置：" + LogHelper.State.Peek(), true);
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
                if (LogHelper.HasException)
                {
                    LogHelper.AddProblemFilesAndReset(filePath);
                }
                else
                {
                    File.Delete(filePath);
                }
            }
        }

        public override JobType JobType
        {
            get { return JobType.GeneratingForm; }
        }
    }
}
