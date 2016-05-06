using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using System.Reflection;
using System.Runtime.InteropServices;

using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;
using PdfSharp;

using Statistics.Criterion.Dose;
using Statistics.Criterion.KV;
using Statistics.Log;
using Statistics.Office.Excel;

namespace Statistics
{
    /// <summary>
    /// 数字(Range.NumberFormatlocal 属性)
    /// 常规：Range.NumberFormatlocal = "G/通用格式"
    /// 数值：Range.NumberFormatlocal = "0.000_" --保留小数位数为3  (此处“_”表示：留下一个与下一个字符同等宽度的空格)
    ///       Range.NumberFormatlocal = "0" --不要小数
    ///       Range.NumberFormatlo cal = "#,##0.000" --保留小数位数为3，并使用千位分隔符
    /// 货币：Range.NumberFormatlocal = "$#,##0.000"
    /// 百分比：Range.NumberFormatlocal = "0.000%"
    /// 分数：Range.NumberFormatlocal = "# ?/?"
    /// 科学计数：Range.NumberFormatlocal = "0.00E+00"
    /// 文本：Range.NumberFormatlocal = "@"
    /// 特殊：Range.NumberFormatlocal = "000000"---邮政编码
    ///       Range.NumberFormatlocal = "[DBNum1]G/通用格式"---中文小写数字
    ///       Range.NumberFormatlocal = "[DBNum2]G/通用格式"---中文大写数字
    ///       Range.NumberFormatlocal = "[DBNum2][$RMB]G/通用格式"---人民币大写
    /// 对齐
    /// 水平对齐：Range.HorizontalAlignment = etHAlignCenter  ---居中
    /// 垂 直对齐：Range.VerticalAlignment = etVAlignCenter---居中
    /// 是否自动换行：Range.WrapText = True
    /// 是否缩小字体填充：Range.ShrinkToFit = True
    /// 是否合并单元格：Range.MergeCells = False
    /// 文字竖排：Range.Orientation = etVertical
    /// 文字倾斜度数：Range.Orientation = 45 -----倾斜45度
    /// 字体（Font对象）
    /// 字体名称：Font.Name = "华文行楷"
    /// 字形：Font.FontStyle = "常规"
    /// 字号：Font.Size = "10"
    /// 下划线：Font.Strikethrough = True; Font.Underline = etUnderlineStyleDouble ---双下划线
    /// 上标：Font.Superscript = True
    /// 下 标：Font.SubScript = True
    /// 删除线：Font.OutlineFont = True
    /// 边框（Borders对象）
    /// Borders.Item(etEdgeTop)：上边框
    /// Borders.Item(etEdgeLeft)：左边框
    /// Borders.Item (etEdgeRight)：右边框
    /// Borders.Item(etEdgeBottom)：下边框
    /// Borders.Item(etDiagonalDown) ：左上--右下边框
    /// Borders.Item(etDiagonalUp)：左下--右上边框
    /// Border.LineStyle = etContinuous 线条样式
    /// </summary>
    public class ExcelUtility
    {
        private MSExcel._Application _excelApp = null;
        public static Regex poRegex = new Regex(@"^[A-Z]+[0-9]+$");

        private MSExcel._Workbook excelDoc = null;
        private string _desiredName = "";

        public string path = "";
        public static int docNumber = 0;
        public static object doNotSaveChanges = MSExcel.XlSaveAction.xlDoNotSaveChanges;

        public ExcelUtility(string Path, out bool success)
        {
            path = Path;
            if (path.ToLower().EndsWith(@".xls") || path.ToLower().EndsWith(@".xlsx"))
            {
                if (File.Exists(path))
                {
                     Init(ref excelDoc, out success);
                }
                else
                {
                    LogHelper.AddException(@"文件不存在" + Environment.NewLine + path, true);
                    success = false;
                }
            }
            else
            {
                LogHelper.AddLog(@"异常25", @"文件不是常见的excel文档类型", true);
                success = false;
            }
        }

        private void Init(ref MSExcel._Workbook wb, out bool success)
        {
            try
            {
                _excelApp = new MSExcel.Application();
                //int[] oldproc = GetPIDs(@"EXCEL");
                object oMissiong = System.Reflection.Missing.Value;

                wb = _excelApp.Workbooks.Open(path, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                success = true;
                //获取Excel App的句柄
                hwnd = new IntPtr(ExcelApp.Hwnd);
                //通过Windows API获取Excel进程ID
                GetWindowThreadProcessId(hwnd, out pid);
                _desiredName = path;
            }
            catch (System.Exception ex)
            {
                LogHelper.AddLog(@"异常24", ex.Message, true);
                success = false;
            }
        }

        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern int OleSetClipboard(IDataObject pDataObj);

        #region Process&Thread

        [DllImport(@"User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        //函数原型；DWORD GetWindowThreadProcessld(HWND hwnd，LPDWORD lpdwProcessld);
        //参数：hWnd:窗口句柄
        //参数：lpdwProcessld:接收进程标识的32位值的地址。如果这个参数不为NULL，GetWindwThreadProcessld将进程标识拷贝到这个32位值中，否则不拷贝
        //返回值：返回值为创建窗口的线程标识。

        [DllImport(@"kernel32.dll")]
        public static extern int OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);
        //函数原型：HANDLE OpenProcess(DWORD dwDesiredAccess,BOOL bInheritHandle,DWORD dwProcessId);
        //参数：dwDesiredAccess：访问权限。
        //参数：bInheritHandle：继承标志。
        //参数：dwProcessId：进程ID。

        public const int PROCESS_ALL_ACCESS = 0x1F0FFF;
        public const int PROCESS_VM_READ = 0x0010;
        public const int PROCESS_VM_WRITE = 0x0020;
        
        //定义句柄变量
        public IntPtr hwnd;

        //定义进程ID变量
        public int pid = -1;

        #endregion

        #region BasicOperation
        /// <summary>
        /// 获取指定字符串位置单元格的range变量
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="position"></param>
        /// <param name="success"></param>
        /// <returns></returns>
        public MSExcel.Range GetRange(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            if (_excelDoc != null)
            {
                if (position.IsValid)
                {
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    MSExcel.Range _excelRge = (MSExcel.Range)_excelSht.Cells.get_Range(position.PositionString, position.PositionString);
                    return _excelRge;
                }
                else
                {
                    LogHelper.AddLog(@"异常26", @"读取数据时传入了错误的位置坐标：" + position.PositionString, true);
                    return null;
                }
            }
            else
            {
                LogHelper.AddLog(@"异常27", @"文件没有正常打开，无法读取数据", true);
                return null;
            }
        }
        /// <summary>
        /// 获取指定坐标位置单元格的range变量
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="success"></param>
        /// <returns></returns>
        public MSExcel.Range GetRange(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position1, ExcelPosition position2)
        {
            if (_excelDoc != null)
            {
                MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                MSExcel.Range _excelRge = (MSExcel.Range)_excelSht.Cells.get_Range(position1.PositionString, position2.PositionString);
                return _excelRge;
            }
            else
            {
                Log.LogHelper.AddLog(@"异常28", @"文件没有正常打开，无法读取数据", true);
                return null;
            }
        }

        public string GetText(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            try
            {
                MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
                return _excelRge.Text.ToString().Trim();
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 向某个坐标位置写入文本
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colomnIndex"></param>
        /// <param name="wValue"></param>
        /// <param name="success"></param>
        public void WriteValue(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string wValue)
        {
            WriteValue(_excelDoc, sheetIndex, position, wValue, "@");
        }
        /// <summary>
        /// 向某个坐标位置写入特定格式的文本
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colomnIndex"></param>
        /// <param name="wValue"></param>
        /// <param name="numberFormat"></param>
        /// <param name="success"></param>
        public void WriteValue(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string wValue, string numberFormat)
        {
            if (_excelDoc != null)
            {
                try
                {
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    _excelSht.Cells[position.RowIndex, position.ColumnIndex] = wValue;
                    MSExcel.Range _excelRge = _excelSht.get_Range(position.PositionString);
                    _excelRge.NumberFormatLocal = numberFormat;
                    return;
                }
                catch (Exception ex)
                {
                    LogHelper.AddLog(@"异常230", ex.Message, true);
                    LogHelper.AddLog(@"异常230", "位置：" + position.PositionString, true);
                    return;
                }
            }
            else
            {
                LogHelper.AddLog(@"异常35", @"文件没有正常打开，无法读取数据", true);
                return;
            }
        }
        /// <summary>
        /// 向某个坐标位置写入公式
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colomnIndex"></param>
        /// <param name="wValue"></param>
        /// <param name="success"></param>
        public void WriteFormula(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string wValue)
        {
            if (_excelDoc != null)
            {
                try
                {
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    MSExcel.Range _excelRge = _excelSht.get_Range(position.PositionString);
                    _excelRge.FormulaLocal = wValue;
                }
                catch (Exception ex)
                {
                    LogHelper.AddLog(@"异常36", ex.Message, true);
                    LogHelper.AddLog(@"异常37", "  " + ex.TargetSite.ToString(), true);
                }
            }
            else
            {
                LogHelper.AddLog(@"异常38", @"文件没有正常打开，无法读取数据", true);
            }
        }
        /// <summary>
        /// 从某个位置复制数据，并向其他某个坐标位置粘贴
        /// </summary>
        /// <param name="sourceSheetIndex"></param>
        /// <param name="sourcePosition"></param>
        /// <param name="destinationSheetIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colomnIndex"></param>
        /// <param name="success"></param>
        public void CopyData(MSExcel._Workbook sourceExcelDoc, int sourceSheetIndex, ExcelPosition sourcePosition, MSExcel._Workbook destinationExcelDoc, int destinationSheetIndex, ExcelPosition destinationPosition, out bool success)
        {
            try
            {
                string temp;
                MSExcel.Range _excelRge = GetRange(sourceExcelDoc, sourceSheetIndex, sourcePosition);
                if (_excelRge.Text.ToString().StartsWith(@"#DI") || _excelRge.Value2 == null)
                {
                    temp = "/";
                    success = false;
                }
                else
                {
                    temp = _excelRge.Value2.ToString();
                    success = true;
                }

                WriteValue(destinationExcelDoc, destinationSheetIndex, destinationPosition, temp);
            }
            catch (System.Exception ex)
            {
                LogHelper.AddLog(@"异常39", ex.Message, true);
                LogHelper.AddLog(@"异常40", "  " + ex.TargetSite.ToString(), true);
                success = false;
            }
        }
        /// <summary>
        /// 从某个位置复制数据，并向其他某个坐标位置以特定格式粘贴
        /// </summary>
        /// <param name="sourceSheetIndex"></param>
        /// <param name="sourcePosition"></param>
        /// <param name="destinationSheetIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colomnIndex"></param>
        /// <param name="numberFormat"></param>
        /// <param name="success"></param>
        public void CopyData(MSExcel._Workbook sourceExcelDoc, int sourceSheetIndex, ExcelPosition sourcePosition, MSExcel._Workbook destinationExcelDoc, int destinationSheetIndex, ExcelPosition destinationPosition, string numberFormat, out bool success)
        {
            try
            {
                string temp;
                MSExcel.Range _excelRge = GetRange(sourceExcelDoc, sourceSheetIndex, sourcePosition);
                if (_excelRge.Text.ToString().StartsWith(@"#DI") || _excelRge.Value2 == null)
                {
                    temp = "/";
                    success = false;
                }
                else
                {
                    temp = GetText(sourceExcelDoc, sourceSheetIndex, sourcePosition);
                    success = true;
                }

                WriteValue(destinationExcelDoc, destinationSheetIndex, destinationPosition, temp, numberFormat);
            }
            catch (System.Exception ex)
            {
                LogHelper.AddLog(@"异常41", ex.Message, true);
                LogHelper.AddLog(@"异常42", "  " + ex.TargetSite.ToString(), true);
                success = false;
            }
        }
        /// <summary>
        /// 智能提取混合内容
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="newRow"></param>
        /// <param name="newCol"></param>
        /// <param name="title"></param>
        /// <param name="success"></param>
        /// <returns></returns>
        public string GetMergeContent(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition startPosition, ExcelPosition endPosition, string title, out bool success)
        {
            string temp_text = GetText(_excelDoc, sheetIndex, startPosition).Replace(@":", "：").Replace(@" ", "");
            
            if (temp_text.StartsWith(title))
            {
                if (temp_text.EndsWith(title))
                {
                    temp_text = GetText(_excelDoc, sheetIndex, endPosition);
                    if (temp_text != "")
                    {
                        success = true;
                        return temp_text;
                    }
                    else
                    {
                        success = false;
                        return "/";
                    }
                }
                else
                {
                    temp_text = temp_text.Replace(title, "").Trim();
                    success = true;
                    return temp_text;
                }
            }
            else
            {
                success = false;
                return "";
            }
        }
        /// <summary>
        /// 智能提取混合内容
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="newRow"></param>
        /// <param name="newCol"></param>
        /// <param name="titles"></param>
        /// <param name="success"></param>
        /// <returns></returns>
        public string GetMergeContent(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition startPosition, ExcelPosition endPosition, string[] titles, out bool success)
        {
            string temp_text1 = GetText(_excelDoc, sheetIndex, startPosition).Replace(@":", "：").Replace(@" ", "");
            string temp_text2 = GetText(_excelDoc, sheetIndex, endPosition);
            string title = "";
            if (!temp_text1.Equals(""))
            {
                foreach (string item in titles)
                {
                    if (temp_text1.StartsWith(item))
                    {
                        if (temp_text1.Equals(item))
                        {
                            title = item;
                            if (temp_text2 != "")
                            {
                                success = true;
                                return temp_text2;
                            }
                            else
                            {
                                success = false;
                                return "/";
                            }
                        }
                        else
                        {
                            success = true;
                            return temp_text1.Replace(item, "").Trim();
                        }
                    }
                }
            }
            success = false;
            return "";
        }
        /// <summary>  
        /// 将图片插入到指定的单元格位置。  
        /// 注意：图片必须是绝对物理路径    /// </summary>  
        /// <param name="RangeName">单元格名称，例如：B4</param>  
        /// <param name="PicturePath">要插入图片的绝对路径。</param> 
        public void WriteImage(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string fileName)
        {
            if (_excelDoc != null)
            {
                try
                {
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    MSExcel.Range _excelRge = _excelSht.get_Range(position.PositionString);
                    _excelRge.Select();
                    MSExcel.Pictures pics = (MSExcel.Pictures)_excelSht.Pictures(Missing.Value);
                    pics.Insert(fileName, Missing.Value);
                    //IDataObject data = null;
                    //data.SetData(DataFormats.Bitmap, image);

                    //OleSetClipboard(data);
                    //Clipboard.SetData(DataFormats.Bitmap, image);
                    //_excelSht.Paste();
                }
                catch (Exception ex)
                {
                    LogHelper.AddLog(@"异常30", ex.Message, true);
                    LogHelper.AddLog(@"异常31", "  " + ex.TargetSite.ToString(), true);
                }
            }
            else
            {
                LogHelper.AddLog(@"异常32", @"文件没有正常打开，无法读取数据", true);
            }
        }
        /// <summary> 
        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。
        /// 注意：图片必须是绝对物理路径    
        /// </summary>   
        /// <param name="RangeName">单元格名称，例如：B4</param>  
        /// <param name="PicturePath">要插入图片的绝对路径。</param>  
        /// <param name="PictuteWidth">插入后，图片在Excel中显示的宽度。</param>    
        /// <param name="PictureHeight">插入后，图片在Excel中显示的高度。</param>    
        public void WriteImage(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position, string fileName, float PictuteWidth, float PictureHeight)
        {
            if (_excelDoc != null)
            {
                try
                {
                    float PicLeft, PicTop;
                    MSExcel.Worksheet _excelSht = (MSExcel.Worksheet)_excelDoc.Worksheets[sheetIndex];
                    MSExcel.Range _excelRge = _excelSht.get_Range(position.PositionString);
                    _excelRge.Select();
                    PicLeft = Convert.ToSingle(_excelRge.Left);
                    PicTop = Convert.ToSingle(_excelRge.Top);
                    _excelSht.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, PictuteWidth, PictureHeight);
                }
                catch (Exception ex)
                {
                    LogHelper.AddLog(@"异常30", ex.Message, true);
                    LogHelper.AddLog(@"异常31", "  " + ex.TargetSite.ToString(), true);
                }
            }
            else
            {
                LogHelper.AddLog(@"异常32", @"文件没有正常打开，无法读取数据", true);
            }
        }

        public bool HadNumber(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            double tempDig;
            string text = GetText(_excelDoc, sheetIndex, position);
            if (text != "-2146826281" && double.TryParse(text, out tempDig))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// opt= 1:普通剂量数据页面
        /// opt= 2:普通CT数据页面
        /// opt= 3:统计页面
        /// </summary>
        /// <param name="crit"></param>
        /// <param name="option"></param>
        /// <returns></returns>
        public int GetColumnByCriterion(NormalDoseCriterion crit, int option)
        {
            if (option == 1)
            {
                switch (crit.Voltage)
                {
                    case "60kV":
                        return 4;
                    case "70kV":
                        return 6;
                    case "80kV":
                        return 8;
                    case "100kV":
                        return 10;
                    case "120kV":
                        return 12;
                    default:
                        LogHelper.AddException("获取列索引的规范不是剂量常用的规范", true);
                        return 16;
                }
            }
            else if (option == 2)
            {
                switch (crit.Voltage)
                {
                    case "60kV":
                        return 4;
                    case "80kV":
                        return 6;
                    case "100kV":
                        return 8;
                    case "120kV":
                        return 10;
                    case "140kV":
                        return 12;
                    default:
                        LogHelper.AddException("获取列索引的规范不是CT常用的规范", true);
                        return 16;
                }
            }
            else if (option == 3)
            {
                return crit.Column;
            }
            else
            {
                LogHelper.AddException("获取列索引时输入了不合法的选项", true);
                return 16;
            }
        }

        public int AddNewCriterionItem(NormalDoseCriterion crit, ref Dictionary<NormalDoseCriterion, int> criList, int startIndex, out bool checkClear)
        {
            bool hasValueFlag = false;
            ArrayList availableNumber = new ArrayList();
            //可能插入的位置
            for (int i = startIndex; i < 23; i += 2)
            {
                availableNumber.Add(i);
            }
            //排除已占有的可能，并检查要加入的规范是否为已有的内容
            foreach (KeyValuePair<NormalDoseCriterion, int> item in criList)
            {
                availableNumber.Remove(item.Value);
                if (item.Key == crit)
                {
                    hasValueFlag = true;
                }
            }
            //寻找最靠前的空位插入规范
            if (!hasValueFlag && availableNumber.Count > 0)
            {
                int number = 23;
                foreach (int item in availableNumber)
                {
                    if (item < number)
                    {
                        number = item;
                    }
                }
                criList.Add(crit, number);
                checkClear = true;
                return number;
            }
            checkClear = false;
            return 4;
        }

        public int MergeCriterions(Dictionary<NormalDoseCriterion, int> criList, ref Dictionary<NormalDoseCriterion, int> criListPack, out bool checkClear)
        {
            int startIndex = 4, tempStartIndex;
            checkClear = true;
            foreach (KeyValuePair<NormalDoseCriterion, int> item in criList)
            {
                tempStartIndex = AddNewCriterionItem(item.Key, ref criList, startIndex, out checkClear);
                if (checkClear)
                {
                    startIndex = tempStartIndex;
                }
                else
                {
                    LogHelper.AddException("分析规范数据时发生错误", true);
                    checkClear = false;
                    break;
                }
            }
            //TODO:加2的方法不严谨，对于有过空位的情况会报错
            return startIndex + 2;
        }

        public bool GetCriterion(MSExcel._Workbook _excelDoc, int sheetIndex, int columnIndex, bool writeCriToData, out string criText, out NormalDoseCriterion crit)
        {
            //确定数据的规范
            NormalDoseCriterion dataCri = NormalDoseCriterion.Null;
            string text = ((MSExcel._Worksheet)_excelDoc.Sheets[sheetIndex]).Name;

            if (text != "标准模板")
            {
                if (text == "统计")
                {
                    text = GetText(_excelDoc, sheetIndex, new ExcelPosition(5, columnIndex));
                    if (text.EndsWith("kV"))
                    {
                        text = text.Replace("kV", "").Trim();
                        switch (text)
                        {
                            case "40":
                                dataCri = NormalDoseCriterion.RQR2_40;
                                break;
                            case "50":
                                dataCri = NormalDoseCriterion.RQR3_50;
                                break;
                            case "60":
                                dataCri = NormalDoseCriterion.RQR4_60;
                                break;
                            case "70":
                                dataCri = NormalDoseCriterion.RQR5_70;
                                break;
                            case "80":
                                dataCri = NormalDoseCriterion.RQR6_80;
                                break;
                            case "90":
                                dataCri = NormalDoseCriterion.RQR7_90;
                                break;
                            case "100":
                                dataCri = NormalDoseCriterion.RQR8_100;
                                break;
                            case "120":
                                dataCri = NormalDoseCriterion.RQR9_120;
                                break;
                            case "140":
                                dataCri = NormalDoseCriterion.RQR_140;
                                break;
                            case "150":
                                dataCri = NormalDoseCriterion.RQR10_150;
                                break;
                        }
                    }
                }
                else
                {
                    text = GetText(_excelDoc, sheetIndex, new ExcelPosition(13, columnIndex));
                    if (text.Contains("140"))
                    {
                        dataCri = NormalDoseCriterion.RQR_140;
                    }
                    else
                    {
                        if (text.Contains("kV"))
                        {
                            string text1 = text.Split(new string[] { "kV" }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                            //GetText(_excelDoc, sheetIndex, DataUtility.DataUtility.PositionString(14, columnIndex), out checkClear);
                            if (text1.Contains("."))
                            {
                                text1 = text1.Substring(0, text1.LastIndexOf('.'));
                            }
                            int charindex = -1;
                            for (int i = 0; i < text1.Length; i++)
                            {
                                if (!DataUtility.DataUtility.IsCharNumber(text1[i]))
                                {
                                    charindex = i;
                                }
                            }
                            if (charindex > -1)
                            {
                                text1 = text1.Substring(charindex + 1);
                            }
                            switch (text1)
                            {
                                case "40":
                                    dataCri = NormalDoseCriterion.RQR2_40;
                                    break;
                                case "50":
                                    dataCri = NormalDoseCriterion.RQR3_50;
                                    break;
                                case "54":
                                    dataCri = NormalDoseCriterion.RQR4_60;
                                    //if (text1.Contains("60"))
                                    //{
                                    //    dataCri = Criterion.RQR4_60;
                                    //}
                                    break;
                                case "60":
                                    dataCri = NormalDoseCriterion.RQR4_60;
                                    break;
                                case "65":
                                    dataCri = NormalDoseCriterion.RQR5_70;
                                    //if (text1.Contains("70"))
                                    //{
                                    //    dataCri = Criterion.RQR5_70;
                                    //}
                                    break;
                                case "70":
                                    dataCri = NormalDoseCriterion.RQR5_70;
                                    break;
                                case "76":
                                    dataCri = NormalDoseCriterion.RQR6_80;
                                    //if (text1.Contains("80"))
                                    //{
                                    //    dataCri = Criterion.RQR6_80;
                                    //}
                                    break;
                                case "80":
                                    dataCri = NormalDoseCriterion.RQR6_80;
                                    break;
                                case "90":
                                    dataCri = NormalDoseCriterion.RQR7_90;
                                    break;
                                case "97":
                                    dataCri = NormalDoseCriterion.RQR8_100;
                                    //if (text1.Contains("100"))
                                    //{
                                    //    dataCri = Criterion.RQR8_100;
                                    //}
                                    break;
                                case "100":
                                    dataCri = NormalDoseCriterion.RQR8_100;
                                    break;
                                case "120":
                                    dataCri = NormalDoseCriterion.RQR9_120;
                                    //if (text1.Contains("120"))
                                    //{
                                    //    dataCri = Criterion.RQR9_120;
                                    //}
                                    break;
                                case "140":
                                    dataCri = NormalDoseCriterion.RQR_140;
                                    //if (text1.Contains("140"))
                                    //{
                                    //    dataCri = Criterion.RQR_140;
                                    //}
                                    break;
                                case "150":
                                    dataCri = NormalDoseCriterion.RQR10_150;
                                    break;
                            }
                        }
                    }
                }
            }
            
            criText = text;
            crit = dataCri;
            if (crit != NormalDoseCriterion.Null)
            {
                if (writeCriToData)
                {
                    WriteCriterionToData(_excelDoc, sheetIndex, crit, columnIndex);
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool GetCriterion(MSExcel._Workbook _excelDoc, int sheetIndex, int columnIndex, bool writeCriToData, out string criText, out KVCriterion crit)
        {
            //确定数据的规范
            KVCriterion dataCri = KVCriterion.Null;
            string text = ((MSExcel._Worksheet)_excelDoc.Sheets[sheetIndex]).Name;

            if (text != "标准模板")
            {
                if (text == "统计")
                {
                    text = GetText(_excelDoc, sheetIndex, new ExcelPosition(5, columnIndex));
                    if (text.EndsWith("kV"))
                    {
                        text = text.Replace("kV", "").Trim();
                        switch (text)
                        {
                            case "40":
                                dataCri = KVCriterion.RQR2_40;
                                break;
                            case "50":
                                dataCri = KVCriterion.RQR3_50;
                                break;
                            case "60":
                                dataCri = KVCriterion.RQR4_60;
                                break;
                            case "70":
                                dataCri = KVCriterion.RQR5_70;
                                break;
                            case "80":
                                dataCri = KVCriterion.RQR6_80;
                                break;
                            case "90":
                                dataCri = KVCriterion.RQR7_90;
                                break;
                            case "100":
                                dataCri = KVCriterion.RQR8_100;
                                break;
                            case "120":
                                dataCri = KVCriterion.RQR9_120;
                                break;
                            case "140":
                                dataCri = KVCriterion.RQR_140;
                                break;
                            case "150":
                                dataCri = KVCriterion.RQR10_150;
                                break;
                        }
                    }
                }
                else
                {
                    text = GetText(_excelDoc, sheetIndex, new ExcelPosition(13, columnIndex));
                    if (text.Contains("140"))
                    {
                        dataCri = KVCriterion.RQR_140;
                    }
                    else
                    {
                        if (text.Contains("kV"))
                        {
                            string text1 = text.Split(new string[] { "kV" }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                            if (text1.Contains("."))
                            {
                                text1 = text1.Substring(0, text1.LastIndexOf('.'));
                            }
                            int charindex = -1;
                            for (int i = 0; i < text1.Length; i++)
                            {
                                if (!DataUtility.DataUtility.IsCharNumber(text1[i]))
                                {
                                    charindex = i;
                                }
                            }
                            if (charindex > -1)
                            {
                                text1 = text1.Substring(charindex + 1);
                            }
                            //GetText(_excelDoc, sheetIndex, DataUtility.DataUtility.PositionString(14, columnIndex), out checkClear);
                            switch (text1)
                            {
                                case "40":
                                    dataCri = KVCriterion.RQR2_40;
                                    break;
                                case "50":
                                    dataCri = KVCriterion.RQR3_50;
                                    break;
                                case "60":
                                    dataCri = KVCriterion.RQR4_60;
                                    break;
                                case "70":
                                    dataCri = KVCriterion.RQR5_70;
                                    break;
                                case "80":
                                    dataCri = KVCriterion.RQR6_80;
                                    break;
                                case "90":
                                    dataCri = KVCriterion.RQR7_90;
                                    break;
                                case "100":
                                    dataCri = KVCriterion.RQR8_100;
                                    break;
                                case "120":
                                    dataCri = KVCriterion.RQR9_120;
                                    break;
                                case "140":
                                    dataCri = KVCriterion.RQR_140;
                                    break;
                                case "150":
                                    dataCri = KVCriterion.RQR10_150;
                                    break;
                            }
                        }
                        else
                        {
                            LogHelper.AddException("无法识别的规范", true);
                        }
                    }
                }
            }

            criText = text;
            crit = dataCri;
            if (crit != KVCriterion.Null)
            {
                if (writeCriToData)
                {
                    WriteCriterionToData(_excelDoc, sheetIndex, crit, columnIndex);
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public int GetCriterions(MSExcel._Workbook _excelDoc, int sheetIndex, out Dictionary<NormalDoseCriterion, int> criList)
        {
            //确定数据的规范
            NormalDoseCriterion dataCri = NormalDoseCriterion.Null;
            string text = ((MSExcel._Worksheet)_excelDoc.Sheets[sheetIndex]).Name;
            criList = new Dictionary<NormalDoseCriterion,int>();
            if (text != "标准模板")
            {
                if (text == "统计")
                {
                    for (int i = 4; i < 23; i += 2)
                    {
                        if (GetCriterion(_excelDoc, sheetIndex, i, false, out text, out dataCri))
                        {
                            criList.Add(dataCri, i);
                        }
                    }
                }
                else
                {
                    for (int i = 4; i < 13; i += 2)
                    {
                        if (GetCriterion(_excelDoc, sheetIndex, i, true, out text, out dataCri))
                        {
                            criList.Add(dataCri, i);
                        }
                    }
                }
            }
            
            return criList.Count;
        }

        public void WriteCriterionToStatistic(MSExcel._Workbook _excelDoc, int sheetIndex, NormalDoseCriterion cri, int columnIndex)
        {
            string volText = "";
            string halfLayerText = "";

            switch (cri.Voltage)
            {
                case "40kV":
                    volText = @"40kV";
                    halfLayerText = @"1.42";
                    break;
                case "50kV":
                    volText = @"50kV";
                    halfLayerText = @"1.78";
                    break;
                case "60kV":
                    volText = @"60kV";
                    halfLayerText = @"2.19";
                    break;
                case "70kV":
                    volText = @"70kV";
                    halfLayerText = @"2.58";
                    break;
                case "80kV":
                    volText = @"80kV";
                    halfLayerText = @"3.01";
                    break;
                case "90kV":
                    volText = @"90kV";
                    halfLayerText = @"3.48";
                    break;
                case "100kV":
                    volText = @"100kV";
                    halfLayerText = @"3.97";
                    break;
                case "120kV":
                    volText = @"120kV";
                    halfLayerText = @"5.00";
                    break;
                case "140kV":
                    volText = @"140kV";
                    halfLayerText = "";
                    break;
                case "150kV":
                    volText = @"150kV";
                    halfLayerText = @"6.57";
                    break;
                default:
                    LogHelper.AddException("在统计页写入第" + columnIndex + "列的数据时遇到无法识别的规范：" + cri.Voltage, true);
                    break;
            }
            if (volText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(5, columnIndex), volText);
            }
            if (halfLayerText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(6, columnIndex), halfLayerText);
            }
            WriteValue(_excelDoc, sheetIndex, new ExcelPosition(7, columnIndex), @"校准因子");
            WriteValue(_excelDoc, sheetIndex, new ExcelPosition(7, columnIndex + 1), @"年稳定性");
        }

        public void WriteCriterionToStatistic(MSExcel._Workbook _excelDoc, int sheetIndex, KVCriterion cri, int columnIndex)
        {
            string volText = "";
            string halfLayerText = "";
            string title2 = @"相对固有误差";

            switch (cri.Voltage)
            {
                case "40kV":
                    title2 = @"固有误差";
                    volText = @"40kV";
                    halfLayerText = @"1.42";
                    break;
                case "50kV":
                    title2 = @"固有误差";
                    volText = @"50kV";
                    halfLayerText = @"1.78";
                    break;
                case "60kV":
                    volText = @"60kV";
                    halfLayerText = @"2.19";
                    break;
                case "70kV":
                    volText = @"70kV";
                    halfLayerText = @"2.58";
                    break;
                case "80kV":
                    volText = @"80kV";
                    halfLayerText = @"3.01";
                    break;
                case "90kV":
                    volText = @"90kV";
                    halfLayerText = @"3.48";
                    break;
                case "100kV":
                    volText = @"100kV";
                    halfLayerText = @"3.97";
                    break;
                case "120kV":
                    volText = @"120kV";
                    halfLayerText = @"5.00";
                    break;
                case "140kV":
                    volText = @"140kV";
                    halfLayerText = "";
                    break;
                case "150kV":
                    volText = @"150kV";
                    halfLayerText = @"6.57";
                    break;
                default:
                    LogHelper.AddException("在统计页写入第" + columnIndex + "列的数据时遇到无法识别的规范：" + cri.Voltage, true);
                    break;
            }
            if (volText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(5, columnIndex), volText);
            }
            if (halfLayerText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(6, columnIndex), halfLayerText);
            }
            WriteValue(_excelDoc, sheetIndex, new ExcelPosition(7, columnIndex), @"平均值");
            WriteValue(_excelDoc, sheetIndex, new ExcelPosition(7, columnIndex + 1), title2);
            WriteValue(_excelDoc, sheetIndex, new ExcelPosition(7, columnIndex + 2), @"年稳定性");
        }

        public void WriteCriterionToData(MSExcel._Workbook _excelDoc, int sheetIndex, NormalDoseCriterion cri, int columnIndex)
        {
            string criText = "";
            string volText = "";
            string halfLayerText = "";

            switch (cri.Voltage)
            {
                case "40kV":
                    criText = @"RQR2(40kV)";
                    volText = @"40kV";
                    halfLayerText = @"1.42";
                    break;
                case "50kV":
                    criText = @"RQR3(50kV)";
                    volText = @"50kV";
                    halfLayerText = @"1.78";
                    break;
                case "60kV":
                    criText = @"RQR4(54kV,250mA,1s)";
                    volText = @"60kV";
                    halfLayerText = @"2.19";
                    break;
                case "70kV":
                    criText = @"RQR5(70kV)";
                    volText = @"70kV";
                    halfLayerText = @"2.58";
                    break;
                case "80kV":
                    criText = @"RQR6(76kV,100mA,1s)";
                    volText = @"80kV";
                    halfLayerText = @"3.01";
                    break;
                case "90kV":
                    criText = @"RQR7(90kV)";
                    volText = @"90kV";
                    halfLayerText = @"3.48";
                    break;
                case "100kV":
                    criText = @"RQR8(97kV,63mA,1s)";
                    volText = @"100kV";
                    halfLayerText = @"3.97";
                    break;
                case "120kV":
                    criText = @"RQR9(120kV,50mA,1s)";
                    volText = @"120kV";
                    halfLayerText = @"5.00";
                    break;
                case "140kV":
                    criText = @"(140kV,32mA,1s)";
                    volText = @"140kV";
                    halfLayerText = "";
                    break;
                case "150kV":
                    criText = @"RQR10(150kV)";
                    volText = @"150kV";
                    halfLayerText = @"6.57";
                    break;
                default:
                    LogHelper.AddException("在数据页写入第" + columnIndex + "列的数据时遇到无法识别的规范：" + cri.Voltage, true);
                    break;
            }
            if (criText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(13, columnIndex), criText);
            }
            if (volText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(14, columnIndex), volText);
            }
            if (halfLayerText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(15, columnIndex), halfLayerText);
            }
        }

        public void WriteCriterionToData(MSExcel._Workbook _excelDoc, int sheetIndex, KVCriterion cri, int columnIndex)
        {
            string volText = cri.Voltage;
            string ppvText = cri.PPV;

            if (ppvText != "")
            {
                WriteValue(_excelDoc, sheetIndex, new ExcelPosition(14, columnIndex), ppvText);
            }
        }
        #endregion

        #region Operation
        /// <summary>
        /// 初始化统计页，规范每页标签，返回统计页的索引
        /// 输出文档中所有出现过的规范集和上下限位置
        /// </summary>
        /// <param name="_excelDoc"></param>
        /// <param name="exSheets"></param>
        /// <param name="firstTime"></param>
        /// <param name="criList"></param>
        /// <param name="limitPosition"></param>
        /// <param name="checkClear"></param>
        /// <returns></returns>
        public int InitialStatisticSheet(MSExcel._Workbook _excelDoc, out Dictionary<int, string> exSheets, out bool firstTime, out Dictionary<NormalDoseCriterion,int> criList, out int limitPosition, out bool checkClear)
        {
            int stateIndex = -1, countIndex;
            MSExcel.Range rr = null;
            string certId;
            List<string> sheetsName = new List<string>();

            Dictionary<NormalDoseCriterion, int> newCriList = new Dictionary<NormalDoseCriterion, int>();

            limitPosition = 24;//limitPosition = 14;
            firstTime = false;
            exSheets = new Dictionary<int, string>();
            criList = new Dictionary<NormalDoseCriterion, int>(); 

            //第一次循环：寻找统计页，统计sheetsName
            foreach (MSExcel.Worksheet item in _excelDoc.Sheets)
            {
                if (item.Name == "统计")
                {
                    stateIndex = item.Index;
                }
                sheetsName.Add(item.Name);
            }
            //检测是否找到了统计页
            if (stateIndex == -1)
            {
                //没找到统计页，创建
                MSExcel._Worksheet m_Sheet = (MSExcel._Worksheet)(_excelDoc.Worksheets.Add(_excelDoc.Sheets[1], Type.Missing, Type.Missing, Type.Missing));
                m_Sheet.Name = "统计";
                stateIndex = m_Sheet.Index;
                firstTime = true;

                if (stateIndex == -1)
                {
                    checkClear = false;
                    return -1;
                }
            }
            else
            {
                //找到统计页后提取已有规范，更新上下限位置
                GetCriterions(_excelDoc, stateIndex, out criList);
                limitPosition = 4 + 2 * criList.Count;
            }
            //第二次循环：获取信息，并规范每页的标签
            foreach (MSExcel.Worksheet item in _excelDoc.Sheets)
            {
                if (item.Name == "统计")
                {
                    stateIndex = item.Index;
                }
                else
                {
                    //规范sheet标签名为证书编号
                    rr = GetRange(_excelDoc, item.Index, new ExcelPosition("L2"));
                    certId = rr.Text.ToString().Trim();
                    if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                    {
                        //有规范的证书号
                        if (item.Name != certId)
                        {
                            sheetsName.Remove(item.Name);
                            if (!sheetsName.Contains(certId))
                            {
                                item.Name = certId;
                            }
                            else
                            {
                                countIndex = 1;
                                while (sheetsName.Contains(certId + "-" + countIndex))
                                {
                                    countIndex++;
                                }
                                item.Name = certId + "-" + countIndex;
                                LogHelper.AddException(@"有两个数据页包含了相同的证书编号", true);
                            }
                            sheetsName.Add(item.Name);
                        }
                        exSheets.Add(item.Index, certId);
                        //把每个数据页出现的规范都加入统计页的规范集中
                        if (GetCriterions(_excelDoc, stateIndex, out newCriList) > 0)
                        {
                            limitPosition = MergeCriterions(newCriList, ref criList, out checkClear);
                            firstTime = true;
                        }
                    }
                    else
                    {
                        //无规范的证书号
                        if (!item.Name.Contains(@"标准模板") && GetText(_excelDoc, item.Index, new ExcelPosition("A4")).StartsWith(@"送校单位"))
                        {
                            //有记录不包含规范的证书编号
                            LogHelper.AddException(@"该文档有实验数据不包含证书编号", true);
                        }
                    }
                    rr = null;
                }
            }
            checkClear = true;
            return stateIndex;
        }
        /// <summary>
        /// 初始化统计页，规范每页标签，返回统计页的索引
        /// </summary>
        /// <param name="_excelDoc"></param>
        /// <param name="exSheets"></param>
        /// <param name="firstTime"></param>
        /// <param name="checkClear"></param>
        /// <returns></returns>
        public int InitialStatisticSheet(MSExcel._Workbook _excelDoc, out Dictionary<int, string> exSheets, out bool firstTime, out bool checkClear)
        {
            int stateIndex = -1, countIndex;
            MSExcel.Range rr = null;
            string certId;
            List<string> sheetsName = new List<string>();

            //Dictionary<DoseCriterion, int> newCriList = new Dictionary<DoseCriterion, int>();

            firstTime = false;
            exSheets = new Dictionary<int, string>();

            //第一次循环：寻找统计页，统计sheetsName
            foreach (MSExcel.Worksheet item in _excelDoc.Sheets)
            {
                if (item.Name == "统计")
                {
                    stateIndex = item.Index;
                }
                sheetsName.Add(item.Name);
            }
            //检测是否找到了统计页
            if (stateIndex == -1)
            {
                //没找到统计页，创建
                MSExcel._Worksheet m_Sheet = (MSExcel._Worksheet)(_excelDoc.Worksheets.Add(_excelDoc.Sheets[1], Type.Missing, Type.Missing, Type.Missing));
                m_Sheet.Name = "统计";
                stateIndex = m_Sheet.Index;
                firstTime = true;

                if (stateIndex == -1)
                {
                    checkClear = false;
                    return -1;
                }
            }
            else
            {
                //找到统计页后提取已有规范，更新上下限位置
                //GetCriterions(_excelDoc, stateIndex, out criList);
                //limitPosition = 4 + 2 * criList.Count;
            }
            //第二次循环：获取信息，并规范每页的标签
            foreach (MSExcel.Worksheet item in _excelDoc.Sheets)
            {
                if (item.Name == "统计")
                {
                    stateIndex = item.Index;
                }
                else
                {
                    //规范sheet标签名为证书编号
                    rr = GetRange(_excelDoc, item.Index, new ExcelPosition("L2"));
                    certId = rr.Text.ToString().Trim();
                    if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                    {
                        //有规范的证书号
                        if (item.Name != certId)
                        {
                            sheetsName.Remove(item.Name);
                            if (!sheetsName.Contains(certId))
                            {
                                item.Name = certId;
                            }
                            else
                            {
                                countIndex = 1;
                                while (sheetsName.Contains(certId + "-" + countIndex))
                                {
                                    countIndex++;
                                }
                                item.Name = certId + "-" + countIndex;
                                LogHelper.AddException(@"有两个数据页包含了相同的证书编号", true);
                            }
                            sheetsName.Add(item.Name);
                        }
                        exSheets.Add(item.Index, certId);
                        //把每个数据页出现的规范都加入统计也的规范集中
                        //if (GetCriterions(_excelDoc, stateIndex, out newCriList) > 0)
                        //{
                        //    limitPosition = MergeCriterions(newCriList, ref criList, out checkClear);
                        //    firstTime = true;
                        //}
                    }
                    else
                    {
                        //无规范的证书号
                        if (!item.Name.Contains(@"标准模板") && GetText(_excelDoc, item.Index, new ExcelPosition("A4")).StartsWith(@"送校单位"))
                        {
                            //有记录不包含规范的证书编号
                            LogHelper.AddException(@"该文档有实验数据不包含证书编号", true);
                        }
                    }
                    rr = null;
                }
            }
            checkClear = true;
            
            return stateIndex;
        }
        /// <summary>
        /// 写入固定内容
        /// </summary>
        /// <param name="pattern"></param>
        public void WriteStateTitle(MSExcel._Workbook _excelDoc, int pattern, int stateIndex, int limitPosition)
        {
            switch (pattern)
            {
                case 0:
                    //设置列宽
                    MSExcel._Worksheet WSDose = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGDose = (MSExcel.Range)WSDose.Columns["A:M", System.Type.Missing];
                    RGDose.ColumnWidth = 10;
                    RGDose = (MSExcel.Range)WSDose.Columns["N:O", System.Type.Missing];
                    RGDose.ColumnWidth = 15;
                    RGDose = (MSExcel.Range)WSDose.Columns["C:C", System.Type.Missing];
                    RGDose.ColumnWidth = 25;
                    //设置格式
                    RGDose = (MSExcel.Range)WSDose.Columns["A:C", System.Type.Missing];
                    RGDose.NumberFormatLocal = "@";
                    //设置字号
                    RGDose = (MSExcel.Range)WSDose.Columns["A:O", System.Type.Missing];
                    RGDose.Font.Size = 11;
                    RGDose = (MSExcel.Range)WSDose.Cells[1, 7];
                    RGDose.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");
                    //写入每个规范对应的管电压和半值层
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");
                    //foreach (KeyValuePair<Criterion, int> item in criList)
                    //{
                    //    if (item.Key != Criterion.Null)
                    //    {
                    //        WriteCriterionToStatistic(_excelDoc, stateIndex, item.Key, item.Value);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value, @"校准因子", out checkClear);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value + 1, @"年稳定性", out checkClear);
                    //    }
                    //}
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR2_40, NormalDoseCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR3_50, NormalDoseCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR4_60, NormalDoseCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR5_70, NormalDoseCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR6_80, NormalDoseCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR7_90, NormalDoseCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR8_100, NormalDoseCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR9_120, NormalDoseCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR_140, NormalDoseCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column);
                    
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
                case 1:
                    //设置列宽
                    MSExcel._Worksheet WSCT = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGCT = (MSExcel.Range)WSCT.Columns["A:M", System.Type.Missing];
                    RGCT.ColumnWidth = 10;
                    RGCT = (MSExcel.Range)WSCT.Columns["N:O", System.Type.Missing];
                    RGCT.ColumnWidth = 15;
                    RGCT = (MSExcel.Range)WSCT.Columns["C:C", System.Type.Missing];
                    RGCT.ColumnWidth = 25;
                    //设置格式
                    RGCT = (MSExcel.Range)WSCT.Columns["A:C", System.Type.Missing];
                    RGCT.NumberFormatLocal = "@";
                    //设置字号
                    RGCT = (MSExcel.Range)WSCT.Columns["A:O", System.Type.Missing];
                    RGCT.Font.Size = 11;
                    RGCT = (MSExcel.Range)WSCT.Cells[1, 7];
                    RGCT.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");
                    //foreach (KeyValuePair<Criterion,int> item in criList)
                    //{
                    //    if (item.Key != Criterion.Null)
                    //    {
                    //        WriteCriterionToStatistic(_excelDoc, stateIndex, item.Key, item.Value);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value, @"校准因子", out checkClear);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value + 1, @"年稳定性", out checkClear);
                    //    }
                    //}
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR2_40, NormalDoseCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR3_50, NormalDoseCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR4_60, NormalDoseCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR5_70, NormalDoseCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR6_80, NormalDoseCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR7_90, NormalDoseCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR8_100, NormalDoseCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR9_120, NormalDoseCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR_140, NormalDoseCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");;
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
                case 2:
                    //设置列宽
                    MSExcel._Worksheet WSKV = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGKV = (MSExcel.Range)WSKV.Columns["A:M", System.Type.Missing];
                    RGKV.ColumnWidth = 10;
                    RGKV = (MSExcel.Range)WSKV.Columns["N:O", System.Type.Missing];
                    RGKV.ColumnWidth = 15;
                    RGKV = (MSExcel.Range)WSKV.Columns["C:C", System.Type.Missing];
                    RGKV.ColumnWidth = 25;
                    //设置格式
                    RGKV = (MSExcel.Range)WSKV.Columns["A:C", System.Type.Missing];
                    RGKV.NumberFormatLocal = "@";
                    //设置字号
                    RGKV = (MSExcel.Range)WSKV.Columns["A:O", System.Type.Missing];
                    RGKV.Font.Size = 11;
                    RGKV = (MSExcel.Range)WSKV.Cells[1, 7];
                    RGKV.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");
                    //写入每个规范对应的管电压和半值层
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");

                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR2_40, KVCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR3_50, KVCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR4_60, KVCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR5_70, KVCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR6_80, KVCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR7_90, KVCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR8_100, KVCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR9_120, KVCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR_140, KVCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR10_150, KVCriterion.RQR10_150.Column);
                    
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
            }
        }
        /// <summary>
        /// 写入固定内容
        /// </summary>
        /// <param name="pattern"></param>
        public void WriteStateTitle(MSExcel._Workbook _excelDoc, int pattern, int stateIndex, Dictionary<NormalDoseCriterion, int> criList, int limitPosition)
        {
            switch (pattern)
            {
                case 0:
                    //设置列宽
                    MSExcel._Worksheet WSDose = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGDose = (MSExcel.Range)WSDose.Columns["A:M", System.Type.Missing];
                    RGDose.ColumnWidth = 10;
                    RGDose = (MSExcel.Range)WSDose.Columns["N:O", System.Type.Missing];
                    RGDose.ColumnWidth = 15;
                    RGDose = (MSExcel.Range)WSDose.Columns["C:C", System.Type.Missing];
                    RGDose.ColumnWidth = 25;
                    //设置格式
                    RGDose = (MSExcel.Range)WSDose.Columns["A:C", System.Type.Missing];
                    RGDose.NumberFormatLocal = "@";
                    //设置字号
                    RGDose = (MSExcel.Range)WSDose.Columns["A:O", System.Type.Missing];
                    RGDose.Font.Size = 11;
                    RGDose = (MSExcel.Range)WSDose.Cells[1, 7];
                    RGDose.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");
                    //写入每个规范对应的管电压和半值层
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");
                    //foreach (KeyValuePair<Criterion, int> item in criList)
                    //{
                    //    if (item.Key != Criterion.Null)
                    //    {
                    //        WriteCriterionToStatistic(_excelDoc, stateIndex, item.Key, item.Value);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value, @"校准因子", out checkClear);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value + 1, @"年稳定性", out checkClear);
                    //    }
                    //}
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR2_40, NormalDoseCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR3_50, NormalDoseCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR4_60, NormalDoseCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR5_70, NormalDoseCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR6_80, NormalDoseCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR7_90, NormalDoseCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR8_100, NormalDoseCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR9_120, NormalDoseCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR_140, NormalDoseCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
                case 1:
                    //设置列宽
                    MSExcel._Worksheet WSCT = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGCT = (MSExcel.Range)WSCT.Columns["A:M", System.Type.Missing];
                    RGCT.ColumnWidth = 10;
                    RGCT = (MSExcel.Range)WSCT.Columns["N:O", System.Type.Missing];
                    RGCT.ColumnWidth = 15;
                    RGCT = (MSExcel.Range)WSCT.Columns["C:C", System.Type.Missing];
                    RGCT.ColumnWidth = 25;
                    //设置格式
                    RGCT = (MSExcel.Range)WSCT.Columns["A:C", System.Type.Missing];
                    RGCT.NumberFormatLocal = "@";
                    //设置字号
                    RGCT = (MSExcel.Range)WSCT.Columns["A:O", System.Type.Missing];
                    RGCT.Font.Size = 11;
                    RGCT = (MSExcel.Range)WSCT.Cells[1, 7];
                    RGCT.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");
                    //foreach (KeyValuePair<Criterion,int> item in criList)
                    //{
                    //    if (item.Key != Criterion.Null)
                    //    {
                    //        WriteCriterionToStatistic(_excelDoc, stateIndex, item.Key, item.Value);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value, @"校准因子", out checkClear);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value + 1, @"年稳定性", out checkClear);
                    //    }
                    //}
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR2_40, NormalDoseCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR3_50, NormalDoseCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR4_60, NormalDoseCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR5_70, NormalDoseCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR6_80, NormalDoseCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR7_90, NormalDoseCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR8_100, NormalDoseCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR9_120, NormalDoseCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR_140, NormalDoseCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号"); ;
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
                case 2:
                    //设置列宽
                    MSExcel._Worksheet WSKV = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGKV = (MSExcel.Range)WSKV.Columns["A:M", System.Type.Missing];
                    RGKV.ColumnWidth = 10;
                    RGKV = (MSExcel.Range)WSKV.Columns["N:O", System.Type.Missing];
                    RGKV.ColumnWidth = 15;
                    RGKV = (MSExcel.Range)WSKV.Columns["C:C", System.Type.Missing];
                    RGKV.ColumnWidth = 25;
                    //设置格式
                    RGKV = (MSExcel.Range)WSKV.Columns["A:C", System.Type.Missing];
                    RGKV.NumberFormatLocal = "@";
                    //设置字号
                    RGKV = (MSExcel.Range)WSKV.Columns["A:O", System.Type.Missing];
                    RGKV.Font.Size = 11;
                    RGKV = (MSExcel.Range)WSKV.Cells[1, 7];
                    RGKV.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");
                    //写入每个规范对应的管电压和半值层
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");

                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR2_40, KVCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR3_50, KVCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR4_60, KVCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR5_70, KVCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR6_80, KVCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR7_90, KVCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR8_100, KVCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR9_120, KVCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR_140, KVCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR10_150, KVCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
            }
        }
        /// <summary>
        /// 写入固定内容
        /// </summary>
        /// <param name="pattern"></param>
        public void WriteStateTitle(MSExcel._Workbook _excelDoc, int pattern, int stateIndex, Dictionary<KVCriterion, int> criList, int limitPosition)
        {
            switch (pattern)
            {
                case 0:
                    //设置列宽
                    MSExcel._Worksheet WSDose = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGDose = (MSExcel.Range)WSDose.Columns["A:M", System.Type.Missing];
                    RGDose.ColumnWidth = 10;
                    RGDose = (MSExcel.Range)WSDose.Columns["N:O", System.Type.Missing];
                    RGDose.ColumnWidth = 15;
                    RGDose = (MSExcel.Range)WSDose.Columns["C:C", System.Type.Missing];
                    RGDose.ColumnWidth = 25;
                    //设置格式
                    RGDose = (MSExcel.Range)WSDose.Columns["A:C", System.Type.Missing];
                    RGDose.NumberFormatLocal = "@";
                    //设置字号
                    RGDose = (MSExcel.Range)WSDose.Columns["A:O", System.Type.Missing];
                    RGDose.Font.Size = 11;
                    RGDose = (MSExcel.Range)WSDose.Cells[1, 7];
                    RGDose.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");
                    //写入每个规范对应的管电压和半值层
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");
                    //foreach (KeyValuePair<Criterion, int> item in criList)
                    //{
                    //    if (item.Key != Criterion.Null)
                    //    {
                    //        WriteCriterionToStatistic(_excelDoc, stateIndex, item.Key, item.Value);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value, @"校准因子", out checkClear);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value + 1, @"年稳定性", out checkClear);
                    //    }
                    //}
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR2_40, NormalDoseCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR3_50, NormalDoseCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR4_60, NormalDoseCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR5_70, NormalDoseCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR6_80, NormalDoseCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR7_90, NormalDoseCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR8_100, NormalDoseCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR9_120, NormalDoseCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR_140, NormalDoseCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
                case 1:
                    //设置列宽
                    MSExcel._Worksheet WSCT = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGCT = (MSExcel.Range)WSCT.Columns["A:M", System.Type.Missing];
                    RGCT.ColumnWidth = 10;
                    RGCT = (MSExcel.Range)WSCT.Columns["N:O", System.Type.Missing];
                    RGCT.ColumnWidth = 15;
                    RGCT = (MSExcel.Range)WSCT.Columns["C:C", System.Type.Missing];
                    RGCT.ColumnWidth = 25;
                    //设置格式
                    RGCT = (MSExcel.Range)WSCT.Columns["A:C", System.Type.Missing];
                    RGCT.NumberFormatLocal = "@";
                    //设置字号
                    RGCT = (MSExcel.Range)WSCT.Columns["A:O", System.Type.Missing];
                    RGCT.Font.Size = 11;
                    RGCT = (MSExcel.Range)WSCT.Cells[1, 7];
                    RGCT.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");
                    //foreach (KeyValuePair<Criterion,int> item in criList)
                    //{
                    //    if (item.Key != Criterion.Null)
                    //    {
                    //        WriteCriterionToStatistic(_excelDoc, stateIndex, item.Key, item.Value);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value, @"校准因子", out checkClear);
                    //        WriteValue(_excelDoc, stateIndex, 7, item.Value + 1, @"年稳定性", out checkClear);
                    //    }
                    //}
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR2_40, NormalDoseCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR3_50, NormalDoseCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR4_60, NormalDoseCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR5_70, NormalDoseCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR6_80, NormalDoseCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR7_90, NormalDoseCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR8_100, NormalDoseCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR9_120, NormalDoseCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR_140, NormalDoseCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, NormalDoseCriterion.RQR10_150, NormalDoseCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号"); ;
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
                case 2:
                    //设置列宽
                    MSExcel._Worksheet WSKV = (MSExcel._Worksheet)_excelDoc.Worksheets[stateIndex];
                    MSExcel.Range RGKV = (MSExcel.Range)WSKV.Columns["A:M", System.Type.Missing];
                    RGKV.ColumnWidth = 10;
                    RGKV = (MSExcel.Range)WSKV.Columns["N:O", System.Type.Missing];
                    RGKV.ColumnWidth = 15;
                    RGKV = (MSExcel.Range)WSKV.Columns["C:C", System.Type.Missing];
                    RGKV.ColumnWidth = 25;
                    //设置格式
                    RGKV = (MSExcel.Range)WSKV.Columns["A:C", System.Type.Missing];
                    RGKV.NumberFormatLocal = "@";
                    //设置字号
                    RGKV = (MSExcel.Range)WSKV.Columns["A:O", System.Type.Missing];
                    RGKV.Font.Size = 11;
                    RGKV = (MSExcel.Range)WSKV.Cells[1, 7];
                    RGKV.Font.Size = 15;
                    //逐行写入内容标签
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(1, 7), @"稳定性统计");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(2, 1), @"送校单位：");

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(3, 1), @"仪器名称：");
                    //写入每个规范对应的管电压和半值层
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(5, 1), @"管电压");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(6, 1), @"半值层（mmAl）");

                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR2_40, KVCriterion.RQR2_40.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR3_50, KVCriterion.RQR3_50.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR4_60, KVCriterion.RQR4_60.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR5_70, KVCriterion.RQR5_70.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR6_80, KVCriterion.RQR6_80.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR7_90, KVCriterion.RQR7_90.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR8_100, KVCriterion.RQR8_100.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR9_120, KVCriterion.RQR9_120.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR_140, KVCriterion.RQR_140.Column);
                    WriteCriterionToStatistic(_excelDoc, stateIndex, KVCriterion.RQR10_150, KVCriterion.RQR10_150.Column);

                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 1), @"序号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 2), @"证书编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, 3), @"仪器编号");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition), @"稳定性误差上限");
                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(7, limitPosition + 1), @"稳定性误差下限");
                    break;
            }
        }
        /// <summary>
        /// 对特定的某些单元格计算特定算法的平均值
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="columnId"></param>
        /// <param name="rowIndex"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="success"></param>
        public void FormulaAverageOneColumn(MSExcel._Workbook _excelDoc, int sheetIndex, int columnId, int rowIndex, int startRow, int endRow, out bool success)
        {
            int digNumber = 0;
            for (int i = startRow; i <= endRow; i++)
            {
                if (HadNumber(_excelDoc, sheetIndex, new ExcelPosition(i, columnId)))
                {
                    digNumber++;
                }
            }
            success = true;
            if (digNumber > 1)
            {
                try
                {
                    MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, new ExcelPosition(rowIndex, columnId));
                    _excelRge.Formula = "=STDEV(" + new ExcelPosition(startRow, columnId) + ":" + new ExcelPosition(endRow, columnId) + ")/AVERAGE(" + new ExcelPosition(startRow, columnId) + ":" + new ExcelPosition(endRow, columnId) + ")";
                    _excelRge.NumberFormatLocal = "0.0%";
                }
                catch (System.Exception ex)
                {
                    success = false;
                    Log.LogHelper.AddLog(@"异常43", ex.Message, true);
                    Log.LogHelper.AddLog(@"异常44", "  " + ex.TargetSite.ToString(), true);
                }
            }
        }
        /// <summary>
        /// 计算长期稳定性
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="range"></param>
        /// <param name="success"></param>
        public void StablizeOneCTDose(MSExcel._Workbook _excelDoc, int sheetIndex, int row, int col, out string range, out bool success)
        {
            try
            {
                ExcelPosition pos2 = new ExcelPosition(row, col - 1);
                ExcelPosition pos1 = new ExcelPosition(row - 1, col - 1);
                ExcelPosition pos = new ExcelPosition(row, col);
                success = true;
                MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, pos);
                if (HadNumber(_excelDoc, sheetIndex, pos2) && HadNumber(_excelDoc, sheetIndex, pos1))
                {
                    _excelRge.Formula = "=(" + pos1.PositionString + "-" + pos2.PositionString + ")/" + pos2.PositionString;
                    _excelRge.NumberFormatLocal = "0.0%";
                }
                if (_excelRge != null && _excelRge.Value2 != null)
                {
                    range = _excelRge.Value2.ToString();
                }
                else
                {
                    range = GetText(_excelDoc, sheetIndex, pos);
                }
            }
            catch (System.Exception ex)
            {
                success = false;
                range = null;
                Log.LogHelper.AddLog(@"异常45", ex.Message, true);
                Log.LogHelper.AddLog(@"异常46", "  " + ex.TargetSite.ToString(), true);
            }
        }
        /// <summary>
        /// 计算长期稳定性
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="range"></param>
        /// <param name="success"></param>
        public void StablizeOneKV(MSExcel._Workbook _excelDoc, int sheetIndex, KVCriterion crit, int row, int col, out string range, out bool success)
        {
            try
            {
                ExcelPosition pos1 = new ExcelPosition(row - 1, col - 2);
                ExcelPosition pos2 = new ExcelPosition(row, col - 2);
                ExcelPosition pos = new ExcelPosition(row, col);
                MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, pos);
                if (HadNumber(_excelDoc, sheetIndex, pos2) && HadNumber(_excelDoc, sheetIndex, pos1))
                {
                    if (crit.Index > 2)
                    {
                        _excelRge.Formula = "=(" + pos2.PositionString + "-" + pos1.PositionString + ")/" + pos1.PositionString;
                    }
                    else
                    {
                        _excelRge.Formula = "=" + pos1.PositionString + "-" + pos2.PositionString;
                    }
                    _excelRge.NumberFormatLocal = "0.0%";
                    _excelDoc.Save();
                }
                if (_excelRge != null && _excelRge.Value2 != null)
                {
                    range = _excelRge.Value2.ToString();
                }
                else
                {
                    range = GetText(_excelDoc, sheetIndex, pos);
                }
                success = true;
            }
            catch (System.Exception ex)
            {
                success = false;
                range = null;
                Log.LogHelper.AddLog(@"异常45", ex.Message, true);
                Log.LogHelper.AddLog(@"异常46", "  " + ex.TargetSite.ToString(), true);
            }
        }

        public string ErrorCalculator(MSExcel._Workbook _excelDoc, int sheetIndex, ExcelPosition position)
        {
            MSExcel.Range _excelRge = GetRange(_excelDoc, sheetIndex, position);
            if (_excelRge != null && _excelRge.Value2 != null)
            {
                return _excelRge.Value2.ToString();
            }
            else
            {
                return GetText(_excelDoc, sheetIndex, position);
            }
        }

        public string GetInsNumber(MSExcel._Workbook _excelDoc, int sourceIndex, out bool checkClear)
        {
            string strMacSerial = "", strSensorSerial = "";

            strMacSerial = "主机_" + GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(5, 7), new ExcelPosition(5, 8), new string[] { @"主机编号：", @"编号：" }, out checkClear);
            strSensorSerial = GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(5, 11), new ExcelPosition(5, 12), new string[] { @"探测器编号：", "电离室号：", "探测器号：" }, out checkClear).Trim();
            if (strSensorSerial.Replace(@"/", "").Trim() != "")
            {
                strSensorSerial = "_探测器_" + strSensorSerial.Trim();
            }
            else
            {
                strSensorSerial = "";
            }
            return strMacSerial + strSensorSerial;
        }

        public string GetInsNumber(MSExcel._Workbook _excelDoc, int sourceIndex, out string strMacSerial, out string strSensorSerial, out bool checkClear)
        {
            strMacSerial = "主机_" + GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(5, 7), new ExcelPosition(5, 8), new string[] { @"主机编号：", @"编号：" }, out checkClear);
            strSensorSerial = GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(5, 11), new ExcelPosition(5, 12), new string[] { @"探测器编号：", "电离室号：", "探测器号：" }, out checkClear).Trim();
            if (strSensorSerial.Replace(@"/", " ").Trim() != "")
            {
                strSensorSerial = "_探测器_" + strSensorSerial.Replace(@"/", " ").Trim();
            }
            else
            {
                strSensorSerial = "";
            }
            return strMacSerial + strSensorSerial;
        }

        public string GenerateFileName(MSExcel._Workbook _excelDoc, int sourceIndex, out bool Perfect)
        {
            bool checkClear = true;
            string strCompany = GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(4, 1), new ExcelPosition(4, 2), new string[] { @"送校单位：", @"单位名称：" }, out checkClear);
            string strType = GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(5, 5), new ExcelPosition(5, 6), @"型号：", out checkClear);
            string insNumber = GetInsNumber(_excelDoc, sourceIndex, out checkClear);
            Perfect = ((strCompany.Length * strType.Length * insNumber.Length) > 0);
            return strCompany + "_" + strType + "_" + insNumber;
        }

        public string GenerateFileName(MSExcel._Workbook _excelDoc, int sourceIndex, out string strCompany, out string strType, out string strMacSerial, out string strSensorSerial, out bool Perfect)
        {
            bool checkClear = true;

            strCompany = GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(4, 1), new ExcelPosition(4, 2), new string[] { @"送校单位：", @"单位名称：" }, out checkClear);
            strType = GetMergeContent(_excelDoc, sourceIndex, new ExcelPosition(5, 5), new ExcelPosition(5, 6), @"型号：", out checkClear);
            string insNumber = GetInsNumber(_excelDoc, sourceIndex, out strMacSerial, out strSensorSerial, out checkClear);
            Perfect = ((strCompany.Length * strType.Length * insNumber.Length) > 0);
            return strCompany + "_" + strType + "_" + insNumber;
        }

        public void CopyKVOneYearData(MSExcel._Workbook _excelDoc, int sourceIndex, int stateIndex, int destiLine, string certId, int limitPosition, Dictionary<KVCriterion, int> criList, bool report, out string insNumber)
        {
            bool checkClear;
            int countIndex = destiLine - 7;
            //原数据位置
            int dataline1 = 20;
            int dataline2 = 21;
            //编号
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 1), countIndex.ToString(), "@");
            //证书编号
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 2), certId);
            //仪器编号
            insNumber = GetInsNumber(_excelDoc, sourceIndex, out checkClear);

            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 3), insNumber);
            //稳定性上下限
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, limitPosition), "0.05");
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, limitPosition + 1), "-0.05");
            
            //清空旧数据
            MSExcel._Worksheet ws = (MSExcel._Worksheet)_excelDoc.Sheets[stateIndex];
            for (int i = 4; i < limitPosition; i++)
            {
                ws.Cells[destiLine, i] = "";
            }
            //提取原数据
            string criText = "";
            string rangeText = "";
            double tempDig = 0.0;
            KVCriterion crit = KVCriterion.Null;
            MSExcel.Range rr = null;
            for (int i = 4; i < 13; i += 2)
            {
                if (GetCriterion(_excelDoc, sourceIndex, i, true, out criText, out crit) && criList.ContainsKey(crit))
                {
                    rr = GetRange(_excelDoc, sourceIndex, new ExcelPosition(dataline1, i));
                    if (rr != null && rr.Value2 != null)
                    {
                        rangeText = rr.Value2.ToString();
                        if (rangeText != "-2146826281" && double.TryParse(rangeText, out tempDig))
                        {
                            //平均值
                            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, criList[crit]), tempDig.ToString(), "0.000");
                            if (crit.Index > 2)
                            {
                                dataline2 = 22;
                            }
                            else
                            {
                                dataline2 = 21;
                            }
                            rr = GetRange(_excelDoc, sourceIndex, new ExcelPosition(dataline2, i));
                            if (rr != null && rr.Value2 != null)
                            {
                                rangeText = rr.Value2.ToString();
                                if (rangeText != "-2146826281" && double.TryParse(rangeText, out tempDig))
                                {
                                    //(相对)固有误差
                                    WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, criList[crit] + 1), tempDig.ToString(), "0.000");
                                    //if (crit.Index > 2)
                                    //{
                                    //    //误差限与(相对)固有误差的差
                                    //    //WriteFormula(_excelDoc, stateIndex, destiLine, criList[crit] + 2, "=0.02-" + DataUtility.DataUtility.PositionString(destiLine, criList[crit] + 1), out checkClear);
                                    //}
                                    //else
                                    //{
                                    //    //误差限与(相对)固有误差的差
                                    //    //WriteFormula(_excelDoc, stateIndex, destiLine, criList[crit] + 2, "=1-" + DataUtility.DataUtility.PositionString(destiLine, criList[crit] + 1), out checkClear);
                                    //}
                                }
                                else if (report)
                                {
                                    Log.LogHelper.AddDataError("KV数据第" + sourceIndex + "页第" + dataline2 + "行第" + i + "列的数据不可识别或不存在，已跳过", true);
                                }
                            }
                        }
                        else if (report)
                        {
                            Log.LogHelper.AddDataError("KV数据第" + sourceIndex + "页第" + dataline1 + "行第" + i + "列的数据不可识别或不存在，已跳过", true);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 把一年数据复制到统计页
        /// </summary>
        /// <param name="_excelDoc">excel文档</param>
        /// <param name="sourceIndex">数据来源sheet的索引</param>
        /// <param name="stateIndex">统计页sheet的索引</param>
        /// <param name="destiLine">统计页sheet中当前应写入的行索引</param>
        /// <param name="certId">数据来源sheet的证书编号</param>
        /// <param name="insNumber">数据来源sheet的数据生成的仪器编号</param>
        public void CopyDoseOneYearData(MSExcel._Workbook _excelDoc, int sourceIndex, int stateIndex, int destiLine, string certId, int limitPosition, Dictionary<NormalDoseCriterion, int> criList, out string insNumber)
        {
            bool checkClear;
            int countIndex = destiLine - 7;
            int dataline = 0;
            //编号
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 1), countIndex.ToString(), "@");
            //证书编号
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 2), certId);
            //仪器编号
            insNumber = GetInsNumber(_excelDoc, sourceIndex, out checkClear);

            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 3), insNumber);
            //稳定性上下限
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, limitPosition), "0.05");
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, limitPosition + 1), "-0.05");
            //判断原数据位置
            if (GetRange(_excelDoc, sourceIndex, new ExcelPosition("A24")).Text.ToString().StartsWith(@"K"))
            {
                dataline = 24;
            }
            else if (GetRange(_excelDoc, sourceIndex, new ExcelPosition("A25")).Text.ToString().StartsWith(@"K"))
            {
                dataline = 25;
            }
            else
            {
                Log.LogHelper.AddException(@"无法提取实验数据，数据行不在24行或25行", true);
                return;
            }
            //清空旧数据
            MSExcel._Worksheet ws = (MSExcel._Worksheet)_excelDoc.Sheets[stateIndex];
            for (int i = 4; i < limitPosition; i++)
            {
                ws.Cells[destiLine, i] = "";
            }
            //提取原数据
            string criText = "";
            string rangeText = "";
            double tempDig;
            NormalDoseCriterion crit = NormalDoseCriterion.Null;
            MSExcel.Range rr = null;
            for (int i = 4; i < 13; i += 2)
            {
                if (GetCriterion(_excelDoc, sourceIndex, i, true, out criText, out crit) && criList.ContainsKey(crit))
                {
                    rr = GetRange(_excelDoc, sourceIndex, new ExcelPosition(dataline, i));
                    if (rr != null)
                    {
                        rangeText = rr.Value2.ToString();
                        if (rangeText != "-2146826281" && double.TryParse(rangeText, out tempDig))
                        {
                            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, criList[crit]), tempDig.ToString(), "0.000");
                        }
                    }
                }
            }
            //提取原数据
            //CopyData(_excelDoc, sourceIndex, "D" + dataline.ToString(), _excelDoc, stateIndex, destiLine, 4, "0.000_ ", out checkClear);
            //CopyData(_excelDoc, sourceIndex, "F" + dataline.ToString(), _excelDoc, stateIndex, destiLine, 6, "0.000_ ", out checkClear);
            //CopyData(_excelDoc, sourceIndex, "H" + dataline.ToString(), _excelDoc, stateIndex, destiLine, 8, "0.000_ ", out checkClear);
            //CopyData(_excelDoc, sourceIndex, "J" + dataline.ToString(), _excelDoc, stateIndex, destiLine, 10, "0.000_ ", out checkClear);
            //CopyData(_excelDoc, sourceIndex, "L" + dataline.ToString(), _excelDoc, stateIndex, destiLine, 12, "0.000_ ", out checkClear);
        }
        
        public void CopyCTOneYearData(MSExcel._Workbook _excelDoc, int sourceIndex, int stateIndex, int destiLine, string certId, int limitPosition, Dictionary<NormalDoseCriterion, int> criList, out string insNumber)
        {
            bool checkClear;
            int countIndex = destiLine - 7;
            int dataline = 0;
            //编号
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 1), countIndex.ToString(), "@");
            //证书编号
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 2), certId);
            //仪器编号
            insNumber = GetInsNumber(_excelDoc, sourceIndex, out checkClear);

            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, 3), insNumber);
            //稳定性上下限
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, limitPosition), "0.05");
            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, limitPosition + 1), "-0.05");
            //判断原数据位置
            if (GetRange(_excelDoc, sourceIndex, new ExcelPosition("A24")).Text.ToString().StartsWith(@"K"))
            {
                dataline = 24;
            }
            else if (GetRange(_excelDoc, sourceIndex, new ExcelPosition("A25")).Text.ToString().StartsWith(@"K"))
            {
                dataline = 25;
            }
            else
            {
                Log.LogHelper.AddException(@"无法提取实验数据，数据行不在24行或25行", true);
                return;
            }
            //清空旧数据
            MSExcel._Worksheet ws = (MSExcel._Worksheet)_excelDoc.Sheets[stateIndex];
            for (int i = 4; i < limitPosition; i++)
            {
                ws.Cells[destiLine, i] = "";
            }
            //判断单位是否含有cm，以进行倍数调整
            string unitText = GetText(_excelDoc, sourceIndex, new ExcelPosition("M12"));
            double multipler = 1;
            if (unitText.ToLower().EndsWith("cm"))
            {
                multipler = 10;
            }
            else if (unitText.ToLower().EndsWith("mm"))
            {
                multipler = 100;
            }
            else if (unitText.ToLower().EndsWith("m"))
            {
                multipler = 0.1;
            }
            //提取原数据
            string criText = "";
            string rangeText = "";
            double tempDig;
            NormalDoseCriterion crit = NormalDoseCriterion.Null;
            MSExcel.Range rr = null;
            for (int i = 4; i < 13; i += 2)
            {
                if (GetCriterion(_excelDoc, sourceIndex, i, true, out criText, out crit) && criList.ContainsKey(crit))
                {
                    //CopyData(_excelDoc, sourceIndex, DataUtility.DataUtility.PositionString(dataline, i), _excelDoc, stateIndex, destiLine, criList[crit], "0.000_ ", out checkClear);
                    rr = GetRange(_excelDoc, sourceIndex, new ExcelPosition(dataline, i));
                    if (rr != null)
                    {
                        rangeText = rr.Value2.ToString();
                        if (rangeText != "-2146826281" && double.TryParse(rangeText, out tempDig))
                        {
                            tempDig *= multipler;
                            WriteValue(_excelDoc, stateIndex, new ExcelPosition(destiLine, criList[crit]), tempDig.ToString(), "0.000");
                        }
                    }
                }
            }
        }

        public bool StatisticsKVOneColumn(MSExcel._Workbook _excelDoc, int stateIndex, int lineIndex, int columnId, Dictionary<int, string> cert, string profile, bool doAverage, bool report, int newline)
        {
            //TODO: 加入KV的统计逻辑
            bool checkClear, pass = true;
            double dig1, dig2;
            string digStr1, digStr2, critStr;
            KVCriterion crit = KVCriterion.Null;
            if (GetCriterion(_excelDoc, stateIndex, columnId, false, out critStr, out crit))
            {
                for (int i = 8; i < lineIndex; i++)
                {
                    digStr1 = ErrorCalculator(_excelDoc, stateIndex, new ExcelPosition(i, columnId + 1));
                    if (report && (newline < 0 || i == newline))
                    {
                        double.TryParse(digStr1, out dig1);
                        if (dig1 > crit.Threshold1)
                        {
                            if (crit.Index > 2)
                            {
                                Log.LogHelper.AddDataError(cert[i] + "在" + profile + "规范下，固有误差过大：" + dig1.ToString("0.00"), true);
                            }
                            else if (crit.Index > 0)
                            {
                                Log.LogHelper.AddDataError(cert[i] + "在" + profile + "规范下，相对固有误差过大：" + dig1.ToString("0.00"), true);
                            }
                            pass = false;
                        }
                    }
                    if (doAverage)
                    {
                        StablizeOneKV(_excelDoc, stateIndex, crit, i, columnId + 2, out digStr2, out checkClear);
                        if (report && (newline < 0 || i == newline))
                        {
                            double.TryParse(digStr2, out dig2);
                            if (Math.Abs(dig2) > crit.Threshold2)
                            {
                                Log.LogHelper.AddDataError(cert[i] + "在" + profile + "规范下年稳定性超差：" + dig2.ToString("0.0%"), true);
                                pass = false;
                            }
                        }
                    }
                }
            }
            return pass;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="_excelDoc">excel文档</param>
        /// <param name="stateIndex">统计页的索引</param>
        /// <param name="lineIndex">当前操作行</param>
        /// <param name="columnId">当前操作列</param>
        /// <param name="cert">页索引与证书编号对应关系</param>
        /// <param name="profile">规范字符串</param>
        /// <param name="doAverage">是否统计长期稳定性</param>
        /// <param name="report">是否超差报错</param>
        /// <returns></returns>
        public bool StatisticsCTDoseOneColumn(MSExcel._Workbook _excelDoc, int stateIndex, int lineIndex, int columnId, Dictionary<int, string> cert, string profile, bool doAverage, bool report, int newline)
        {
            bool checkClear, pass = true;
            double dig;
            string digStr;
            FormulaAverageOneColumn(_excelDoc, stateIndex, columnId, lineIndex, 8, lineIndex - 1, out checkClear);
            if (doAverage)
            {
                for (int i = 9; i < lineIndex; i++)
                {
                    StablizeOneCTDose(_excelDoc, stateIndex, i, columnId + 1, out digStr, out checkClear);
                    if (report && (newline < 0 || i == newline) && double.TryParse(digStr, out dig))
                    {
                        if (Math.Abs(dig) >= 0.05)
                        {
                            Log.LogHelper.AddDataError(cert[i] + "在" + profile + "规范下年稳定性超差：" + dig.ToString("0.0%"), true);
                            pass = false;
                        }
                    }
                }
            }
            return pass;
        }

        public string Verification(MSExcel._Workbook _excelDoc, bool fulltest, int pattern, out bool pass)
        {
            string certId = "";
            string standardId = "";
            string fileName = "";
            bool checkClear;
            bool firstTime = true;
            bool sameCom, sameMac;
            string strCompany = "", strMacSerial = "", strSensor = "", strType = "";
            string strCompanynew = "", strMacSerialnew = "", strSensornew = "", strTypenew = "";
            string stdStrCompany = "", stdStrMacSerial = "", stdStrSensor = "", stdStrType = "";
            string stdStrCompanynew = "", stdStrMacSerialnew = "", stdStrSensornew = "", stdStrTypenew = "";

            pass = true;

            //获取信息
            foreach (MSExcel.Worksheet item in _excelDoc.Sheets)
            {
                if (item.Name != "统计" && item.Name != "标准模板")
                {
                    //规范sheet标签名为证书编号
                    certId = GetText(_excelDoc, item.Index, new ExcelPosition("L2"));
                    if (certId.StartsWith(@"20") && (certId.Length == 9 || certId.Length == 10))
                    {
                        if (firstTime)
                        {
                            fileName = GenerateFileName(_excelDoc, item.Index, out stdStrCompany, out stdStrType, out stdStrMacSerial, out stdStrSensor, out checkClear);
                            stdStrMacSerialnew = stdStrMacSerial.ToLower();
                            stdStrTypenew = stdStrType.ToLower();
                            stdStrCompanynew = stdStrCompany.ToLower();
                            if (stdStrSensor != "")
                            {
                                stdStrSensornew = stdStrSensor.ToLower();
                            }
                            standardId = certId;
                            firstTime = false;
                        }
                        else
                        {
                            GenerateFileName(_excelDoc, item.Index, out strCompany, out strType, out strMacSerial, out strSensor, out checkClear);
                            strMacSerialnew = strMacSerial.ToLower();
                            strTypenew = strType.ToLower();
                            strCompanynew = strCompany.ToLower();
                            if (strSensor != "")
                            {
                                strSensornew = strSensor.ToLower();
                                if (strSensornew != stdStrSensornew)
                                {
                                    pass = false;
                                    Log.LogHelper.AddDataError(@"检查校验：数据中包含不同的探测器编号", true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + standardId + ":" + stdStrSensor, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + certId + ":" + strSensor, true);
                                }
                            }
                            else
                            {
                                sameCom = (strCompanynew == stdStrCompanynew);
                                sameMac = (strMacSerialnew == stdStrMacSerialnew);
                                if (sameCom && !sameMac)
                                {
                                    pass = false;
                                    Log.LogHelper.AddDataError(@"检查校验：主机编号不同", true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + standardId + ":" + stdStrMacSerial, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + certId + ":" + strMacSerial, true);
                                }
                                if (!sameCom && sameMac)
                                {
                                    pass = false;
                                    Log.LogHelper.AddDataError(@"检查校验：单位名称不同", true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + standardId + ":" + stdStrCompany, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + certId + ":" + strCompany, true);
                                }
                                if (!sameCom && !sameMac)
                                {
                                    pass = false;
                                    Log.LogHelper.AddDataError(@"检查校验：单位名称和主机编号均不对应", true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + standardId + ":" + stdStrCompany, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + certId + ":" + strCompany, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + standardId + ":" + stdStrMacSerial, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + certId + ":" + strMacSerial, true);
                                }
                                if (strTypenew != stdStrTypenew)
                                {
                                    pass = false;
                                    Log.LogHelper.AddDataError(@"检查校验：仪器名称不同", true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + standardId + ":" + stdStrType, true);
                                    Log.LogHelper.AddLog(@"                检查校验：  " + certId + ":" + strType, true);
                                }
                            }
                        }

                        if (pattern == 0)
                        {
                            string text = "";
                            double dig = 0;
                            bool isDouble = false;
                            for (int columnIndex = 4; columnIndex < 13; columnIndex += 2)
                            {
                                text = GetText(_excelDoc, item.Index, new ExcelPosition(24, columnIndex));
                                isDouble = double.TryParse(text, out dig);
                                if (isDouble)
                                {
                                    if (dig > 2.0 || dig < 0.1)
                                    {
                                        string reportText = "检查校验：" + item.Name + "页" + new ExcelPosition(24, columnIndex).PositionString + "位置出现异常的数据";
                                        if (fulltest)
                                        {
                                            Log.LogHelper.AddException(reportText, true);
                                        }
                                        else
                                        {
                                            Log.LogHelper.AddDataError(reportText, true);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        //无规范的证书号
                        if (!item.Name.Contains(@"标准模板") && GetText(_excelDoc, item.Index, new ExcelPosition("A4")).StartsWith(@"送校单位"))
                        {
                            //有记录不包含规范的证书编号
                            pass = false;
                            Log.LogHelper.AddException(@"检查校验：该文档有实验数据不包含证书编号", true);
                        }
                    }
                }
            }

            if (pass)
            {
                return fileName;
            }
            else
            {
                return "";
            }
        }

        public void SaveAsPDF(MSExcel._Workbook _excelDoc, string targetPath, string tempPath, out bool success)
        {
            object missing = Type.Missing;
            try
            {
                string temp = Util.PathExt.PathChangeDirectory(targetPath, tempPath);
                _excelDoc.ExportAsFixedFormat(MSExcel.XlFixedFormatType.xlTypePDF, temp, MSExcel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, false, missing);
                success = PDFEncrypt(temp, targetPath, true);
            }
            catch (Exception ex)
            {
                Log.LogHelper.AddException(ex.Message, true);
                success = false;
            }
        }

        private bool PDFEncrypt(string filenameSource, string filenameDest, bool deleteOld)
        {
            try
            {
                // Open an existing document. Providing an unrequired password is ignored.
                PdfSharp.Pdf.PdfDocument doc = PdfSharp.Pdf.IO.PdfReader.Open(filenameSource, "password");

                PdfSharp.Pdf.Security.PdfSecuritySettings securitySettings = doc.SecuritySettings;

                // Setting one of the passwords automatically sets the security level to 
                // PdfDocumentSecurityLevel.Encrypted128Bit.
                //securitySettings.UserPassword = "";//DataUtility.DataUtility.EncryptGenerateRandomString(DataUtility.DataUtility.ASCIIMode.DigitalLowerUpper, 8);
                securitySettings.OwnerPassword = "dyjlnim";// DataUtility.DataUtility.EncryptGenerateRandomString(DataUtility.DataUtility.ASCIIMode.DigitalLowerUpper, 8);

                // Don't use 40 bit encryption unless needed for compatibility
                //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = false;
                securitySettings.PermitFullQualityPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitPrint = true;

                // Save the document...
                doc.Save(filenameDest);
                if (deleteOld)
                {
                    File.Delete(filenameSource);
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.LogHelper.AddException("PDF加密时出错：" + ex.Message, true);
                return false;
            }
        }
        
        #endregion

        #region Exit&Disposal

        public void TryClose()
        {
            try
            {
                //正常退出Excel
                //ExcelApp.Quit()后,若Form_KillExcelProcess进程正常结束,那么,Excel进程也会自动结束.
                //ExcelApp.Quit()后,若Form_KillExcelProcess进程被用户手工强制结束,那么,Excel进程不会自动结束.
                excelDoc.Close(doNotSaveChanges, Missing.Value, Missing.Value);
                _excelApp.Quit();
            }
            catch (Exception ex)
            {
                Log.LogHelper.AddLog(@"退出23", "  " + ex.Message, false);
            }
            finally
            {
                DisposeExcel(ref excelDoc);
            }
        }
        /// <summary>
        /// 释放所引用的COM对象。注意：这个过程一定要执行。
        /// </summary>
        public void DisposeExcel(ref MSExcel._Workbook _excelDoc)
        {
            ReleaseObj(_excelDoc);
            ReleaseObj(_excelApp);
            _excelDoc = null;
            _excelApp = null;
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.Threading.Thread.Sleep(1000);

            //强制结束Excel进程
            if (pid > 0) //_excelApp != null && 
            {
                int ExcelProcess = OpenProcess(PROCESS_VM_READ | PROCESS_VM_WRITE, false, pid);
                //判断进程是否仍然存在
                if (ExcelProcess > 0)
                {
                    try
                    {
                        //通过进程ID,找到进程
                        System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(pid);
                        //Kill 进程
                        process.Kill();
                    }
                    catch
                    {
                        //强制结束Excel进程失败,可以记录一下日志.
                        //AddLog(@"退出", "  结束进程失败，异常：" + ex.Message, false);
                        //AddLog(@"退出", "    " + ex.TargetSite, false);
                    }
                }
                else
                {
                    //进程已经结束了
                }
            }
        }

        private void ReleaseObj(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally { o = null; }
        }

        #endregion

        #region Property

        public MSExcel._Application ExcelApp
        {
            get
            {
                return _excelApp;
            }
        }

        public MSExcel._Workbook ExcelWorkbook
        {
            get
            {
                return excelDoc;
            }
        }

        public string DesiredName
        {
            get
            {
                return _desiredName;
            }
            set
            {
                _desiredName = value.Replace(@"/", " ").Replace(@"\", " ").Replace(@":", " ").Replace(@"*", " ").Replace(@"?", " ").Replace(@"<", " ").Replace(@">", " ").Replace(@"|", " ").Replace(@"""", " ");
            }
        }

        #endregion
    }

    public class WordUtility
    {
        private MSWord.ApplicationClass _wordApp = null;

        private MSWord._Document wordDoc = null;
        private string _desiredName = "";
        private Object Nothing = Missing.Value;

        public string path = "";
        public static int docNumber = 0;
        public static object doNotSaveChanges = MSExcel.XlSaveAction.xlDoNotSaveChanges;

        [DllImport(@"User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        //函数原型；DWORD GetWindowThreadProcessld(HWND hwnd，LPDWORD lpdwProcessld);
        //参数：hWnd:窗口句柄
        //参数：lpdwProcessld:接收进程标识的32位值的地址。如果这个参数不为NULL，GetWindwThreadProcessld将进程标识拷贝到这个32位值中，否则不拷贝
        //返回值：返回值为创建窗口的线程标识。

        [DllImport(@"kernel32.dll")]
        public static extern int OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);
        //函数原型：HANDLE OpenProcess(DWORD dwDesiredAccess,BOOL bInheritHandle,DWORD dwProcessId);
        //参数：dwDesiredAccess：访问权限。
        //参数：bInheritHandle：继承标志。
        //参数：dwProcessId：进程ID。

        public const int PROCESS_ALL_ACCESS = 0x1F0FFF;
        public const int PROCESS_VM_READ = 0x0010;
        public const int PROCESS_VM_WRITE = 0x0020;
        
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("User32.dll", EntryPoint = "FindWindowEx")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpClassName, string lpWindowName);
        
        //定义句柄变量
        public IntPtr hwnd;

        //定义进程ID变量
        public int pid = -1;

        public WordUtility(string Path, out bool success)
        {
            path = Path;
            if (path.ToLower().EndsWith(@".doc") || path.ToLower().EndsWith(@".docx"))
            {
                if (File.Exists(path))
                {
                    Init(ref wordDoc, out success);
                }
                else
                {
                    Log.LogHelper.AddException(@"文件不存在" + Environment.NewLine + path, true);
                    success = false;
                }
            }
            else
            {
                Log.LogHelper.AddLog(@"异常25", @"文件不是常见的word文档类型", true);
                success = false;
            }
        }

        private void Init(ref MSWord._Document dc, out bool success)
        {
            try
            {
                string AppID = DateTime.Now.Ticks.ToString();
                _wordApp = new MSWord.ApplicationClass();
                _wordApp.Application.Caption = AppID;
                _wordApp.WindowState = MSWord.WdWindowState.wdWindowStateMinimize;
                _wordApp.Application.Visible = true;
                _wordApp.WindowState = MSWord.WdWindowState.wdWindowStateMinimize;
                
                while (DataUtility.DataUtility.GetProcessIdByWindowTitle(AppID) == Int32.MaxValue)
                {
                    Thread.Sleep(10);
                }
                pid = DataUtility.DataUtility.GetProcessIdByWindowTitle(AppID);
                _wordApp.Application.Visible = false;

                dc = _wordApp.Documents.OpenNoRepairDialog(path);
                success = true;
                _desiredName = path;
            }
            catch (System.Exception ex)
            {
                Log.LogHelper.AddLog(@"异常24", ex.Message, true);
                success = false;
            }
        }

        public string GetText(MSWord._Document myWordDoc, int para)
        {
            string temp = myWordDoc.Paragraphs[para].Range.Text.Trim();
            if (temp == null)
            {
                return "/";
            }
            else
            {
                return temp.Replace(" ", "").Replace("\r", "").Replace("\a", "");
            }
        }

        public void WriteValue(MSWord._Document myWordDoc, object bookmark, string value)
        {
            if (myWordDoc.Bookmarks.Exists(bookmark.ToString()))
            {
                myWordDoc.Bookmarks.get_Item(ref bookmark).Range.Text = value; //WORD  插入文本
            }
        }

        public void WriteValue(MSWord._Document myWordDoc, object bookmark, MSExcel.Range rg)
        {
            object sth = (object)rg.Value;
            if (myWordDoc.Bookmarks.Exists(bookmark.ToString()))
            {
                if (sth == null)
                {
                    myWordDoc.Bookmarks.get_Item(ref bookmark).Range.Text = @"/"; //WORD  插入文本
                }
                else
                {
                    myWordDoc.Bookmarks.get_Item(ref bookmark).Range.Text = sth.ToString(); //WORD  插入文本
                }
            }
            else
            {
                Log.LogHelper.AddException("查找书签失败，书签名：" + bookmark.ToString(), true);
            }
        }

        public void WriteDataValue(MSWord._Document myWordDoc, object bookmark, MSExcel.Range rg, string format)
        {
            object sth = (object)rg.Value;
            if (myWordDoc.Bookmarks.Exists(bookmark.ToString()))
            {
                if (sth == null)
                {
                    myWordDoc.Bookmarks.get_Item(ref bookmark).Range.Text = @"/"; //WORD  插入文本
                }
                else
                {
                    string temp = sth.ToString();
                    if (temp == "-2146826281" || temp == @"/")
                    {
                        myWordDoc.Bookmarks.get_Item(ref bookmark).Range.Text = @"/"; //WORD  插入文本
                    }
                    else
                    {
                        myWordDoc.Bookmarks.get_Item(ref bookmark).Range.Text = string.Format(format, float.Parse(temp)); //WORD  插入文本
                    }
                }
            }
            else
            {
                Log.LogHelper.AddException("查找书签失败，书签名：" + bookmark.ToString(), true);
            }
        }

        public void SaveAsPDF(MSWord._Document myWordDoc, string targetPath, out bool success)
        {
            Object fileformat = MSWord.WdSaveFormat.wdFormatPDF;
            Object missing = System.Reflection.Missing.Value;
            try
            {
                myWordDoc.SaveAs(targetPath, fileformat);
                success = true;
            }
            catch (Exception ex)
            {
                success = false;
                Log.LogHelper.AddException(ex.Message, true);
            }
        }

        #region Exit&Disposal

        public void TryClose()
        {
            try
            {
                //正常退出Excel
                //ExcelApp.Quit()后,若Form_KillExcelProcess进程正常结束,那么,Excel进程也会自动结束.
                //ExcelApp.Quit()后,若Form_KillExcelProcess进程被用户手工强制结束,那么,Excel进程不会自动结束.
                Object sc = MSWord.WdSaveOptions.wdDoNotSaveChanges;
                WordApp.NormalTemplate.Saved = true;
                wordDoc.Saved = true;
                wordDoc.Close(sc, Missing.Value, Missing.Value);
                _wordApp.Quit(ref sc, ref Nothing, ref Nothing);
            }
            catch (Exception ex)
            {
                Log.LogHelper.AddLog(@"退出23", "  " + ex.Message, false);
            }
            finally
            {
                DisposeExcel(ref wordDoc);
            }
        }
        /// <summary>
        /// 释放所引用的COM对象。注意：这个过程一定要执行。
        /// </summary>
        public void DisposeExcel(ref MSWord._Document _wordDoc)
        {
            ReleaseObj(_wordDoc);
            ReleaseObj(_wordApp);
            _wordDoc = null;
            _wordApp = null;
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.Threading.Thread.Sleep(1000);

            //强制结束Excel进程
            if (pid > 0) //_excelApp != null && 
            {
                int ExcelProcess = OpenProcess(PROCESS_VM_READ | PROCESS_VM_WRITE, false, pid);
                //判断进程是否仍然存在
                if (ExcelProcess > 0)
                {
                    try
                    {
                        //通过进程ID,找到进程
                        System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(pid);
                        //Kill 进程
                        process.Kill();
                    }
                    catch
                    {
                        //强制结束Excel进程失败,可以记录一下日志.
                        //AddLog(@"退出", "  结束进程失败，异常：" + ex.Message, false);
                        //AddLog(@"退出", "    " + ex.TargetSite, false);
                    }
                }
                else
                {
                    //进程已经结束了
                }
            }
        }

        private void ReleaseObj(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally { o = null; }
        }

        #endregion

        #region Property

        public MSWord._Application WordApp
        {
            get
            {
                return _wordApp;
            }
        }

        public MSWord._Document WordDocument
        {
            get
            {
                return wordDoc;
            }
        }

        public string DesiredName
        {
            get
            {
                return _desiredName;
            }
            set
            {
                _desiredName = value.Replace(@"/", " ").Replace(@"\", " ").Replace(@":", " ").Replace(@"*", " ").Replace(@"?", " ").Replace(@"<", " ").Replace(@">", " ").Replace(@"|", " ").Replace(@"""", " ");
            }
        }

        #endregion

    }    
}
//class InsertPictureToExcel
//{
//    /// <summary>
//    /// 将图片插入到指定的单元格位置。
//    /// 注意：图片必须是绝对物理路径
//    /// </summary>
//    /// <param name="RangeName">单元格名称，例如：B4</param>
//    /// <param name="PicturePath">要插入图片的绝对路径。</param>
//    public void InsertPicture(string RangeName, string PicturePath)
//    {
//        m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
//        m_objRange.Select();
//        MSExcel.Pictures pics = (MSExcel.Pictures)m_objSheet.Pictures(m_objOpt);
//        pics.Insert(PicturePath, m_objOpt);
//    }

//    /// <summary>
//    /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。
//    /// 注意：图片必须是绝对物理路径
//    /// </summary>
//    /// <param name="RangeName">单元格名称，例如：B4</param>
//    /// <param name="PicturePath">要插入图片的绝对路径。</param>
//    /// <param name="PictuteWidth">插入后，图片在Excel中显示的宽度。</param>
//    /// <param name="PictureHeight">插入后，图片在Excel中显示的高度。</param>
//    public void InsertPicture(string RangeName, string PicturePath, float PictuteWidth, float PictureHeight)
//    {
//        m_objRange = m_objSheet.get_Range(RangeName, m_objOpt);
//        m_objRange.Select();
//        float PicLeft, PicTop;
//        PicLeft = Convert.ToSingle(m_objRange.Left);
//        PicTop = Convert.ToSingle(m_objRange.Top);
//        //参数含义：
//        //图片路径
//        //是否链接到文件
//        //图片插入时是否随文档一起保存
//        //图片在文档中的坐标位置（单位：points）
//        //图片显示的宽度和高度（单位：points）
//        //参数详细信息参见：http://msdn2.microsoft.com/zh-cn/library/aa221765(office.11).aspx
//        m_objSheet.Shapes.AddPicture(PicturePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, PictuteWidth, PictureHeight);
//    }

//    private Excel.Application m_objExcel = null;
//    private Excel.Workbooks m_objBooks = null;
//    private Excel._Workbook m_objBook = null;
//    private Excel.Sheets m_objSheets = null;
//    private Excel._Worksheet m_objSheet = null;
//    private Excel.Range m_objRange = null;
//    private object m_objOpt = System.Reflection.Missing.Value;
//}    
////Set the AppId
//string AppId = ""+DateTime.Now.Ticks(); //A random title

////Create an identity for the app

//this.oWordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
//this.oWordApp.Application.Caption = AppId;
//this.oWordApp.Application.Visible = true;

/////Get the pid by for word application
//this.WordPid = StaticMethods.GetProcessIdByWindowTitle(AppId);

//while ( StaticMethods.GetProcessIdByWindowTitle(AppId) == Int32.MaxValue) //Loop till u get
//{
//    Thread.Sleep(5);
//}

//this.WordPid = StaticMethods.GetProcessIdByWindowTitle(AppId);


/////You canh hide the application afterward            
//this.oWordApp.Application.Visible = false;

//string this.oWordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
//this.oWordApp.Application.Caption = AppId;
//this.oWordApp.Application.Visible = true;
/////Get the pid by 
//this.WordPid = StaticMethods.GetProcessIdByWindowTitle(AppId);

//while ( StaticMethods.GetProcessIdByWindowTitle(AppId) == Int32.MaxValue)
//{
//    Thread.Sleep(5);
//}

//this.WordPid = StaticMethods.GetProcessIdByWindowTitle(AppId);

//this.oWordApp.Application.Visible = false; //You Can hide the application now
