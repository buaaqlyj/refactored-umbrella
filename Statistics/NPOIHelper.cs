using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Runtime.InteropServices;

using Statistics.DataUtility;

using NPOI.DDF;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.OpenXml4Net;
using NPOI.OpenXmlFormats;
using NPOI.POIFS;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF;

namespace Statistics
{
    /// <summary>
    /// 调用第三方Office控件NPOI，主要实现功能：
    /// 一、BasicOperation
    /// 1、GetRange（选取一定位置的cells范围）
    /// - public MSExcel.Range GetRange(int sheetIndex, string position, out bool success)
    /// - public MSExcel.Range GetRange(int sheetIndex, string position1, string position2, out bool success)
    /// - public MSExcel.Range GetRange(int sheetIndex, int row, int col, out bool success)
    /// 
    /// 2、GetText（获取一个单元格的内容）
    /// - public string GetText(int sheetIndex, string position, out bool success)
    /// - public string GetText(int sheetIndex, int row, int col, out bool success)
    /// 
    /// 3、WriteValue（向一个单元格写入数据）
    /// - public void WriteValue(int sheetIndex, int rowIndex, int colomnIndex, string wValue, out bool success)
    /// - public void WriteValue(int sheetIndex, int rowIndex, int colomnIndex, string wValue, string numberFormat,out bool success)
    /// 
    /// 4、WriteFormula（向一个单元格写入公式）
    /// - public void WriteFormula(int sheetIndex, int rowIndex, int colomnIndex, string wValue, out bool success)
    /// 
    /// 5、CopyData（从一个单元格向另一个单元格复制数据）
    /// - public void CopyData(int sourceSheetIndex, string position, int destinationSheetIndex, int rowIndex, int colomnIndex, out bool success)
    /// - public void CopyData(int sourceSheetIndex, string position, int destinationSheetIndex, int rowIndex, int colomnIndex, string numberFormat, out bool success)
    /// 
    /// 6、GetMergeContent（获取连续两个单元格中合并后的文本数据）
    /// - public string GetMergeContent(int sheetIndex, int row, int col, int newRow, int newCol, string title, out bool success)
    /// 
    /// 7、CleanName（去掉文件名中不合法的字符）
    /// - public string CleanName(string name)
    /// 
    /// 8、PositionString（把位置序号转换成文本形式）
    /// - public string PositionString(int row, int col)
    /// 
    /// 二、Operation
    /// 1、InitialStatisticSheet（第一次循环，寻找统计页，统计sheetname；第二次循环，规范sheetname，没有统计页则新建；最后返回统计页序号）
    /// - public int InitialStatisticSheet(out Dictionary<int, string> exSheets, out bool firstTime, out bool checkClear)
    /// 
    /// 2、WriteStateTitle（写入统计页表头，设置列宽，格式）
    /// - public void WriteStateTitle(int pattern, int stateIndex)
    /// 
    /// 3、FormulaAverageOneColumn（在统计页面计算平均值）
    /// - public void FormulaAverageOneColumn(int sheetIndex, string position, string startPosition, string endPosition, out bool success)
    /// - public void FormulaAverageOneColumn(int sheetIndex, int columnId, int rowIndex, int startRow, int endRow, out bool success)
    /// 
    /// 4、StablizeOne（在统计页面统计稳定度）
    /// - public void StablizeOne(int sheetIndex, int row, int col, out MSExcel.Range range, out bool success)
    /// 
    /// 5、GetPureNumber（得到字符串中的纯数字）
    /// - public string GetPureNumber(string text, bool keepSpace)
    /// 
    /// 6、GetInsNumber（返回主机加探测器编号的字符串）
    /// - public string GetInsNumber(int sourceIndex, out bool checkClear)
    /// - public string GetInsNumber(int sourceIndex, out string strMacSerial, out string strSensorSerial, out bool checkClear)
    /// 
    /// 7、PerfectFileName（返回合成的文件名字符串）
    /// - public string PerfectFileName(int sourceIndex, out bool Perfect)
    /// - public string PerfectFileName(int sourceIndex, out string strCompany, out string strType, out string strMacSerial, out string strSensorSerial, out bool Perfect)
    /// 
    /// 8、CopyOneYearData（把某一年的5个数据拷贝到统计页）
    /// - public void CopyOneYearData(int sourceIndex, int stateIndex, int destiLine, string certId, out string insNumber)
    /// 
    /// 9、StatisticsOneColumn（统计一个规范下的年稳定性）
    /// - public void StatisticsOneColumn(int stateIndex, int lineIndex, int columnId, Dictionary<int, string> cert, string profile, bool doAverage, bool report)
    /// 
    /// 10、Verification（检查校验是否有不一致的信息）
    /// - public string Verification(bool addException, int pattern, out bool pass)
    /// 
    /// 11、GetPIDs（）
    /// - public int[] GetPIDs(string name)
    /// 
    /// 12、GetNewProcess（）
    /// - public int[] GetNewProcess(int[] oldProc, int[] newProc)
    /// </summary>
    public class NPOIHelper
    {
        //IWorkbook _sourceWorkbook, _destiWorkbook;
        
        public void Init()
        {
            
        }

        #region IO

        public IWorkbook Open(string FileName, out bool success)
        {
            string path = FileName;
            success = true;
            if (path.ToLower().EndsWith(@".xls") || path.ToLower().EndsWith(@".xlsx"))
            {
                if (File.Exists(path))
                {
                    try
                    {
                        using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                        {
                            return WorkbookFactory.Create(file);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        AddLog(@"异常24", ex.Message, true);
                        success = false;
                    }
                }
                else
                {
                    AddExceptionLog(@"文件不存在" + Environment.NewLine + path, true);
                    //AddLog(@"异常", @"文件不存在" + Environment.NewLine + path, true);
                    success = false;
                }
            }
            else
            {
                AddLog(@"异常25", @"文件不是常见的excel文档类型", true);
                success = false;
            }
            return null;
        }

        public void Save(IWorkbook workbook, string FileName)
        {
            using (FileStream file = new FileStream(FileName, FileMode.Create))
            {
                workbook.Write(file);
                file.Close();
            }
        }
        #endregion

        #region Basic Operation

        public ICell GetRange(IWorkbook workbook, int sheetIndex, int row, int col)
        {
            ISheet s1 = workbook.GetSheetAt(sheetIndex);
            return s1.GetRow(row).GetCell(col);
        }

        public ICell GetRange(IWorkbook workbook, int sheetIndex, int row, int col, out CellType celltype)
        {
            ICell c1 = GetRange(workbook, sheetIndex, row, col);
            celltype = c1.CellType;
            return c1;
        }

        public ICellRange<ICell> GetRange(IWorkbook workbook, int sheetIndex, int frow, int fcol, int srow, int scol)
        {
            ISheet s1 = workbook.GetSheetAt(sheetIndex);
            //TODO: 获取多个单元格的范围
            return null;
        }

        public string GetText(IWorkbook workbook, int sheetIndex, int row, int col)
        {
            return GetRange(workbook, sheetIndex, row, col).StringCellValue;
        }

        public void WriteValue(IWorkbook workbook, int sheetIndex, int rowIndex, int colomnIndex, string wValue, CellType celltype)
        {
            ICell c1 = GetRange(workbook, sheetIndex, rowIndex, colomnIndex);
            c1.SetCellValue(wValue);
            c1.SetCellType(celltype);
        }

        public void WriteFomula(IWorkbook workbook, int sheetIndex, int rowIndex, int colomnIndex, string formula)
        {
            ICell c1 = GetRange(workbook, sheetIndex, rowIndex, colomnIndex);
            c1.SetCellFormula(formula);
        }

        public void CopyData(IWorkbook sourceWorkbook, int sourceSheetIndex, int sourceRowIndex, int sourceColomnIndex, IWorkbook destiWorkbook, int destiSheetIndex, int destiRowIndex, int destiColomnIndex)
        {
            
        }
        #endregion


        #region Log
        public delegate void TextBoxWriteInvoke(string str);
        public static event TextBoxWriteInvoke tbwi;
        public delegate void LogFileWriteInvoke(string str);
        public static event LogFileWriteInvoke lfwi;
        public delegate void AddExceptionDelegate(string ex, bool log);
        public static event AddExceptionDelegate aed;
        public delegate void AddDataErrorDelegate(string ex, bool log);
        public static event AddDataErrorDelegate aded;

        public void AddLog(string pre, string ex, bool sw)
        {
            string temp = @"【" + pre + @"】" + ex;
            if (sw)
            {
                lfwi(temp);
            }
            tbwi(temp + Environment.NewLine);
        }

        public void AddLog(string ex, bool sw)
        {
            if (sw)
            {
                lfwi(ex);
            }
            tbwi(ex + Environment.NewLine);
        }

        public void AddExceptionLog(string ex, bool log)
        {
            aed(ex, log);
        }

        public void AddDataErrorLog(string ex, bool log)
        {
            aded(ex, log);
        }
        #endregion

    }

    public enum OfficeState
    {
        Closed=0,
        Opened=1
    }
}
