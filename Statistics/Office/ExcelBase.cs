﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using System.Reflection;
using System.Runtime.InteropServices;

using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;
using Statistics.Office.Excel;
using PdfSharp;

namespace Statistics.Office
{
    /// <summary>
    /// 承担excel的基础操作
    /// </summary>
    public static class ExcelBase
    {
        private static List<ExcelAppVar> appList = new List<ExcelAppVar>();
        private static Dictionary<int, MSExcel._Workbook> docDic = new Dictionary<int, MSExcel._Workbook>();
        private static Dictionary<int, MSExcel._Application> docAppDic = new Dictionary<int, MSExcel._Application>();

        public static void OpenExcelApp(ref ExcelAppVar ap)
        {
            ap.App = new MSExcel.Application();
            appList.Add(ap);
        }

        public static void OpenExcelDoc(string fileName, ref ExcelAppVar apVar, ref MSExcel._Workbook wb)
        {
            wb = apVar.App.Workbooks.Open(fileName);
            //获取Excel App的句柄
            int pid = -1;
            //通过Windows API获取Excel进程ID
            DataUtility.DataUtility.GetWindowThreadProcessId(new IntPtr(apVar.App.Hwnd), out pid);
            apVar.docDic.Add(pid, wb);
            docAppDic.Add(pid, apVar.App);
        }

        public static void TryCloseDoc(ref MSExcel._Workbook wb)
        {

        }

        public static void TryCloseApp(ref MSExcel._Application ap)
        {

        }

        #region GetRange

        public static MSExcel.Range GetRange(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            return GetRange(_excelDoc, sheetIndex, position, position);
        }

        public static MSExcel.Range GetRange(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position1, ExcelPosition position2)
        {
            if (_excelDoc != null)
            {
                if (position1.IsValid && position2.IsValid)
                {
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    MSExcel.Range _excelRge = (MSExcel.Range)_excelSht.Cells.get_Range(position1, position2);
                    return _excelRge;
                }
                else
                {
                    Log.LogHelper.AddLog(@"异常26", @"读取数据时传入了错误的位置坐标：" + position1, true);
                    return null;
                }
            }
            else
            {
                Log.LogHelper.AddLog(@"异常27", @"文件没有正常打开，无法读取数据", true);
                return null;
            }
        }
        #endregion

        #region GetText

        public static string GetText(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            try
            {
                MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
                return _excelRge.Text.ToString();
            }
            catch
            {
                return "";
            }
        }

        public static string GetValueText(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
            if (_excelRge != null)
            {
                return _excelRge.Value2.ToString();
            }
            else
            {
                return "";
            }
        }

        public static string GetMergedContent(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position1, ExcelPosition position2, string title)
        {
            return GetMergedContent(_excelDoc, sheetIndex, position1, position2, new string[] { title });
        }

        public static string GetMergedContent(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position1, ExcelPosition position2, string[] titles)
        {
            string temp_text1 = GetText(_excelDoc, sheetIndex, position1).Replace(@":", "：").Replace(@" ", "");
            string temp_text2 = GetText(_excelDoc, sheetIndex, position2).Replace(@":", "：").Replace(@" ", "");

            if (!temp_text1.Equals(""))
            {
                foreach (string item in titles)
                {
                    if (temp_text1.Equals(item))
                    {
                        if (temp_text2 != "")
                        {
                            return temp_text2;
                        }
                        else
                        {
                            return "/";
                        }
                    }
                    else
                    {
                        return temp_text1.Replace(item, "").Trim();
                    }
                }
            }

            string text = temp_text1 + temp_text2;

            foreach (string item in titles)
            {
                if (text.StartsWith(item))
                {
                    text = text.Replace(item, "").Trim();
                    if (text != "")
                    {
                        return text;
                    }
                    else
                    {
                        return "/";
                    }
                }
            }
            return "";
        }
        #endregion

        #region WriteStuff

        public static void WriteValue(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string wValue, string numberFormat, out bool success)
        {
            if (_excelDoc != null)
            {
                try
                {
                    bool checkSta = true;
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    _excelSht.Cells[position.RowIndex, position.ColumnIndex] = wValue;
                    if (!string.IsNullOrWhiteSpace(numberFormat))
                    {
                        MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
                        _excelRge.NumberFormatLocal = numberFormat;
                    }
                    success = checkSta;
                    return;
                }
                catch (Exception ex)
                {
                    success = false;
                    Log.LogHelper.AddLog(@"异常33", ex.Message, true);
                    Log.LogHelper.AddLog(@"异常34", "  " + ex.TargetSite.ToString(), true);
                    return;
                }
            }
            else
            {
                success = false;
                Log.LogHelper.AddLog(@"异常35", @"文件没有正常打开，无法读取数据", true);
                return;
            }
        }

        public static void WriteFormula(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string wValue, out bool success)
        {
            if (_excelDoc != null)
            {
                try
                {
                    MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
                    _excelRge.FormulaLocal = wValue;
                    success = true;
                    return;
                }
                catch (Exception ex)
                {
                    success = false;
                    Log.LogHelper.AddLog(@"异常36", ex.Message, true);
                    Log.LogHelper.AddLog(@"异常37", "  " + ex.TargetSite.ToString(), true);
                    return;
                }
            }
            else
            {
                success = false;
                Log.LogHelper.AddLog(@"异常38", @"文件没有正常打开，无法读取数据", true);
                return;
            }
        }

        public static void WriteImage(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string personPath, float PictuteWidth, float PictureHeight)
        {
            if (_excelDoc != null)
            {
                try
                {
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
                    _excelRge.Select();

                    if (PictuteWidth < 1 || PictureHeight < 1)
                    {
                        MSExcel.Pictures pics = (MSExcel.Pictures)_excelSht.Pictures(Missing.Value);
                        pics.Insert(personPath, Missing.Value);
                    }
                    else
                    {
                        float PicLeft = Convert.ToSingle(_excelRge.Left);
                        float PicTop = Convert.ToSingle(_excelRge.Top);
                        _excelSht.Shapes.AddPicture(personPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, PictuteWidth, PictureHeight);
                    }

                    return;
                }
                catch (Exception ex)
                {
                    Log.LogHelper.AddLog(@"异常130", ex.Message, true);
                    Log.LogHelper.AddLog(@"异常131", "  " + ex.TargetSite.ToString(), true);
                    return;
                }
            }
            else
            {
                Log.LogHelper.AddLog(@"异常32", @"文件没有正常打开，无法读取数据", true);
                return;
            }
        }
        #endregion

    }

    public struct ExcelAppVar
    {
        public MSExcel._Application App;
        public Dictionary<int, MSExcel._Workbook> docDic;
    }
}
