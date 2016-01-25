using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

using Statistics.Office;

namespace Statistics.DataUtility
{
    public static class DataUtility
    {
        #region Log
        public delegate void TextBoxWriteInvoke(string str);
        public static event TextBoxWriteInvoke tbwi;
        public delegate void LogFileWriteInvoke(string str);
        public static event LogFileWriteInvoke lfwi;
        public delegate void AddExceptionDelegate(string ex, bool log);
        public static event AddExceptionDelegate aed;
        public delegate void AddDataErrorDelegate(string ex, bool log);
        public static event AddDataErrorDelegate aded;

        public static void AddLog(string pre, string ex, bool sw)
        {
            string temp = @"【" + pre + @"】" + ex;
            if (sw)
            {
                lfwi(temp);
            }
            tbwi(temp + Environment.NewLine);
        }

        public static void AddLog(string ex, bool sw)
        {
            if (sw)
            {
                lfwi(ex);
            }
            tbwi(ex + Environment.NewLine);
        }

        public static void AddExceptionLog(string ex, bool log)
        {
            aed(ex, log);
        }

        public static void AddDataErrorLog(string ex, bool log)
        {
            aded(ex, log);
        }
        #endregion

        #region Folder

        public static void TryCreatFolder(string folderName)
        {
            DirectoryInfo di = new DirectoryInfo(folderName);
            if (!di.Exists)
            {
                di.Create();
            }
        }

        public static void TryCreatFolders(string fatherFolderName, string[] folderNames)
        {
            foreach (string item in folderNames)
            {
                TryCreatFolder(Path.Combine(fatherFolderName, FileNameCleanName(item)));
            }
        }

        public static void TryDeleteFilesInFolders(string folderName, string keyWord)
        {
            FileInfo[] deleteFiles = (new DirectoryInfo(folderName)).GetFiles(keyWord, SearchOption.AllDirectories);
            if (deleteFiles.Count() > 0)
            {
                foreach (FileInfo fi in deleteFiles)
                {
                    File.Delete(fi.FullName);
                }
            }
        }

        #endregion

        #region Char
        /// <summary>
        /// 判断字符是不是数字
        /// </summary>
        /// <param name="ch">输入的字符</param>
        /// <returns>判断结果</returns>
        public static bool IsCharNumber(char ch)
        {
            if (ch > 47 && ch < 58)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 判断字符是不是数字，并返回所在序号
        /// </summary>
        /// <param name="ch">输入的字符</param>
        /// <param name="index">判断结果是数字时输出该数字；否则输出1</param>
        /// <returns>判断结果</returns>
        public static bool IsCharNumber(char ch, out int index)
        {
            if (ch > 47 && ch < 58)
            {
                index = ch - 48;
                return true;
            }
            else
            {
                index = 1;
                return false;
            }
        }
        /// <summary>
        /// 判断字符是不是大写字母
        /// </summary>
        /// <param name="ch">输入的字符</param>
        /// <returns>判断结果</returns>
        public static bool IsCharUpperLetter(char ch)
        {
            if (ch > 64 && ch < 91)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 判断字符是不是大写字母，并返回所在序号
        /// </summary>
        /// <param name="ch">输入的字符</param>
        /// <param name="index">判断结果是大写字母时输出字母所在序号，从1开始；否则输出1</param>
        /// <returns>判断结果</returns>
        public static bool IsCharUpperLetter(char ch, out int index)
        {
            if (ch > 64 && ch < 91)
            {
                index = ch - 64;
                return true;
            }
            else
            {
                index = 1;
                return false;
            }
        }
        /// <summary>
        /// 判断字符是不是小写字母
        /// </summary>
        /// <param name="ch">输入的字符</param>
        /// <returns>判断结果</returns>
        public static bool IsCharLowerLetter(char ch)
        {
            if (ch > 96 && ch < 123)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 判断字符是不是小写字母，并返回所在序号
        /// </summary>
        /// <param name="ch">输入的字符</param>
        /// <param name="index">判断结果是小写字母时输出字母所在序号，从1开始；否则输出1</param>
        /// <returns>判断结果</returns>
        public static bool IsCharLowerLetter(char ch, out int index)
        {
            if (ch > 96 && ch < 123)
            {
                index = ch - 96;
                return true;
            }
            else
            {
                index = 1;
                return false;
            }
        }

        #endregion

        #region Array

        public static FileInfo[] CombineFileInfoArray(FileInfo[] a1, FileInfo[] a2)
        {
            FileInfo[] array = new FileInfo[a1.Length + a2.Length];
            a1.CopyTo(array, 0);
            a2.CopyTo(array, a1.Length);
            return array;
        }

        #endregion

        #region Path&Filename

        public static string PathCombine(string folder1, string folder2)
        {
            return Path.Combine(folder1, folder2);
        }

        public static string PathCombineFileExtension(string file, string extension)
        {
            if (!extension.StartsWith(@".")) extension = @"." + extension;
            return file + extension;
        }

        public static string PathCombineFolderFileExtension(string folder, string file, string extension)
        {
            return PathCombine(folder, PathCombineFileExtension(file, extension));
        }

        public static string PathCombineClassified(string folder, int pattern)
        {
            string keyword = "";
            switch (pattern)
            {
                case 0:
                    //dose
                    keyword = "剂量";
                    break;
                case 1:
                    //ct
                    keyword = "CT";
                    break;
                case 2:
                    //kv
                    keyword = "KV";
                    break;
            }
            return Path.Combine(folder, keyword);
        }

        public static string PathChangeExtension(string path, string extension)
        {
            return Path.ChangeExtension(path, extension);
        }

        public static string PathChangeDirectory(string path, string directory)
        {
            return Path.Combine(directory, Path.GetFileName(path));
        }

        public static string PathGetFileName(string file)
        {
            return Path.GetFileName(file);
        }
        /// <summary>
        /// 去除名称中含有的非法字符
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string FileNameCleanName(string name)
        {
            return name.Replace(@"/", " ").Replace(@"\", " ").Replace(@":", " ").Replace(@"*", " ").Replace(@"?", " ").Replace(@"<", " ").Replace(@">", " ").Replace(@"|", " ").Replace(@"""", " ");
        }


        #endregion

        #region Sort

        public static bool LargerThan(string cert1, string cert2)
        {
            int i1, i2, i3, i4;
            string[] sa = cert1.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            i1 = Int32.Parse(sa[0]);
            i2 = Int32.Parse(sa[1]);
            sa = cert2.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            i3 = Int32.Parse(sa[0]);
            i4 = Int32.Parse(sa[1]);
            if (i1 > i3)
            {
                return true;
            }
            else if (i1 < i3)
            {
                return false;
            }
            else if (i2 >= i4)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static int sortUnit(int[] array, Dictionary<int, string> data, int low, int high)
        {
            int key = array[low];
            while (low < high)
            {
                /*从后向前搜索比key小的值*/
                while (LargerThan(data[array[high]], data[key]) && high > low) --high;
                /*比key小的放左边*/
                array[low] = array[high];

                /*从前向后搜索比key大的值，比key大的放右边*/
                while (LargerThan(data[key], data[array[low]]) && high > low) ++low;
                /*比key大的放右边*/
                array[high] = array[low];
            }
            /*左边都比key小，右边都比key大。
            //将key放在游标当前位置。
            //此时low等于high
            */
            array[low] = key;
            return high;
        }
        /// <summary>
        /// 快速排序
        /// </summary>
        /// <param name="array"></param>
        /// <param name="low"></param>
        /// <param name="high"></param>
        public static void QuickSort(int[] array, Dictionary<int, string> data, int low, int high)
        {
            if (low > high)
                return;
            /*完成一次单元排序*/
            int index = sortUnit(array, data, low, high);
            /*对左边单元进行排序*/
            QuickSort(array, data, low, index - 1);
            /*对右边单元进行排序*/
            QuickSort(array, data, index + 1, high);
        }

        #endregion

        #region String
        /// <summary>
        /// 比较两个字符串的相似指数
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        public static int CompareStrings(string str1, string str2)
        {
            int rate = 0;
            str1 = str1.Trim().ToLower();
            str2 = str2.Trim().ToLower();
            if (str1 == "" || str2 == "")
            {
                rate = 0;
            }
            else
            {
                double r1 = 0;
                double r2 = 0;
                foreach (char item in str1)
                {
                    if (str2.Contains(item.ToString()))
                    {
                        r2++;
                    }
                }
                r2 /= (double)str2.Length;
                foreach (char item in str2)
                {
                    if (str1.Contains(item.ToString()))
                    {
                        r1++;
                    }
                }
                r1 /= (double)str1.Length;
                rate = (int)(100.0 * r1 * r2);
            }

            return rate;
        }

        public static string GetPureNumber(string text, bool keepSpace)
        {
            string pn = "";
            foreach (char item in text)
            {
                if (item >= '0' && item <= '9')
                {
                    pn = pn + item;
                }
                else if (keepSpace && !pn.EndsWith(@" "))
                {
                    pn = pn + " ";
                }
            }
            return pn.Trim();
        }


        #endregion

        #region Config

        public static string ReadInText(string fileName, string ignoreFlag, string separateFlag)
        {
            string text = "";
            string ttext = "";
            if (ignoreFlag == "")
            {
                ignoreFlag = "#";
            }
            if (separateFlag == "")
            {
                separateFlag = ",";
            }
            try
            {
                using (StreamReader typeSR = new StreamReader(fileName))
                {
                    while (!typeSR.EndOfStream)
                    {
                        ttext = typeSR.ReadLine();
                        if (!ttext.StartsWith(ignoreFlag))
                        {
                            text = text + separateFlag + ttext;
                        }
                    }
                }
                return text;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }
        }
        
        #endregion

        #region Process
        /// <summary>
        /// Returns the name of that process given by that title
        /// </summary>
        /// <param name="AppId">Int32MaxValue returned if it cant be found.</param>
        /// <returns></returns>
        public static int GetProcessIdByWindowTitle(string AppId)
        {
            Process[] P_CESSES = Process.GetProcesses();
            for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
            {
                if (P_CESSES[p_count].MainWindowTitle.Equals(AppId))
                {
                    return P_CESSES[p_count].Id;
                }
            }

            return Int32.MaxValue;
        }

        [DllImport(@"User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        //函数原型；DWORD GetWindowThreadProcessld(HWND hwnd，LPDWORD lpdwProcessld);
        //参数：hWnd:窗口句柄
        //参数：lpdwProcessld:接收进程标识的32位值的地址。如果这个参数不为NULL，GetWindwThreadProcessld将进程标识拷贝到这个32位值中，否则不拷贝
        //返回值：返回值为创建窗口的线程标识。


        #endregion

        #region Encrypt

        public enum ASCIIMode { DigitalLowerUpper = 0, Digital = 1, DigitalLower = 2, DigitalUpper = 3, LowerUpper = 4, Lower = 5, Upper = 6};

        public static char EncryptGenerateRandomChar(ASCIIMode mode)
        {
            //这里下界是0，随机数可以取到，上界应该是75，因为随机数取不到上界，也就是最大74，符合我们的题意
            string str = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ~!@#$%^&*()_+";
            int chaNum = 75;
            switch (mode)
            {
                case ASCIIMode.DigitalLowerUpper:
                    str = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    chaNum = 62;
                    break;
                case ASCIIMode.Digital:
                    str = "1234567890";
                    chaNum = 10;
                    break;
                case ASCIIMode.DigitalLower:
                    str = "1234567890abcdefghijklmnopqrstuvwxyz";
                    chaNum = 36;
                    break;
                case ASCIIMode.DigitalUpper:
                    str = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    chaNum = 36;
                    break;
                case ASCIIMode.LowerUpper:
                    str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    chaNum = 52;
                    break;
                case ASCIIMode.Lower:
                    str = "abcdefghijklmnopqrstuvwxyz";
                    chaNum = 26;
                    break;
                case ASCIIMode.Upper:
                    str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    chaNum = 26;
                    break;
            }

            Random r = new Random();
            return str.Substring(r.Next(0, chaNum), 1)[0];
        }

        public static int EncryptGenerateRandomNumber(int low, int high)
        {
            if (low >= high)
            {
                low = 0;
                high = 100;
            }
            Random r = new Random();
            return r.Next(low, high);
        }

        public static string EncryptGenerateRandomString(ASCIIMode mode, int length)
        {
            if (length < 1)
            {
                length = 8;
            }
            string result = string.Empty;

            for (int i = 0; i < length; i++)
            {
                result += EncryptGenerateRandomChar(mode);
            }

            return result;
        }

        public static string EncryptString(string text)
        {
            MD5 md = MD5.Create();
            string txtMd5 = "";
            byte[] bytes = md.ComputeHash(System.Text.Encoding.UTF8.GetBytes(text));
            foreach (byte b in bytes)
            {
                txtMd5 = txtMd5 + b.ToString();
            }
            return txtMd5.Substring(0, 8);
        }

        #endregion

        #region Backup
        private static Dictionary<string, string> BackupList = new Dictionary<string, string>();
        /// <summary>
        /// 备份文件
        /// </summary>
        /// <param name="tempFolder">存放备份文件的临时文件夹</param>
        /// <param name="file">要备份的原文件</param>
        /// <param name="key">恢复密钥</param>
        /// <returns>是否成功</returns>
        public static bool BackupFile(string tempFolder, string file, out string key)
        {
            if (File.Exists(file))
            {
                try
                {
                    //要备份到的位置
                    string newfile = PathCombine(tempFolder, PathGetFileName(file));
                    //要备份到的位置如果存在同名文件即删除
                    if (File.Exists(newfile))
                    {
                        File.Delete(newfile);
                    }
                    //复制备份
                    File.Copy(file, newfile);
                    //生成key
                    key = EncryptString(newfile);
                    //把key和源文件位置加入备份列表
                    if (BackupList.Keys.Contains(key))
                    {
                        BackupList.Remove(key);
                    }
                    BackupList.Add(key, file);
                    return true;
                }
                catch (Exception ex)
                {
                    key = "";
                    AddExceptionLog("备份文件时出错" + ex, true);
                    return false;
                }
            }
            else
            {
                key = "";
                AddExceptionLog("备份的文件不存在", true);
                return false;
            }
        }
        /// <summary>
        /// 恢复备份的文件
        /// </summary>
        /// <param name="tempFolder">存放备份文件的临时文件夹</param>
        /// <param name="key">恢复密钥</param>
        /// <returns>是否成功</returns>
        public static bool RestoreFile(string tempFolder, string key)
        {
            try
            {
                if (key != "" && BackupList.Keys.Contains(key))
                {
                    //源文件应该在的位置
                    string file = BackupList[key];
                    //备份文件位置
                    string backfile = PathCombine(tempFolder, PathGetFileName(file));
                    if (File.Exists(backfile))
                    {
                        if (File.Exists(file))
                        {
                            if (File.Exists(backfile + ".old"))
                            {
                                File.Delete(backfile + ".old");
                            }
                            File.Move(file, backfile+".old");
                        }
                        File.Copy(backfile, file);
                        return true;
                    }
                    else
                    {
                        AddExceptionLog("备份库中找不到所需的文件", true);
                        return false;
                    }
                }
                else
                {
                    AddExceptionLog("备份记录中找不到提供的恢复密钥", true);
                    return false;
                }
            }
            catch (Exception ex)
            {
                AddExceptionLog("恢复备份文件时出错：" + ex, true);
                return false;
            }
        }
        #endregion

        /// <summary>
        /// 根据坐标生成位置字符串
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string PositionString(int row, int col)
        {
            string pos = "";
            ExcelPosition.ChangeNumberToString(row, col, out pos);
            return pos;
        }
        /// <summary>
        /// 根据坐标生成位置字符串
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string PositionString(int row, int col, out bool success)
        {
            string pos = "";
            ExcelPosition.ChangeNumberToString(row, col, out pos);
            success = true;
            return pos;
        }

        public static string GetLastPartFromName(string text)
        {
            if (text.Contains(@"_"))
            {
                string[] strs = text.Split('_');
                return GetPureNumber(strs[strs.Length - 1], false);
            }
            return text;
        }

        public static string PathRightFileName(string folder, string file, string extension, string tag)
        {
            if (tag == "") tag = "-new";
            tag = FileNameCleanName(tag);
            if (!folder.EndsWith(@"\")) folder = folder + @"\";
            if (!extension.StartsWith(@".")) extension = @"." + FileNameCleanName(extension);
            file = FileNameCleanName(file);
            while (File.Exists(folder + file + extension))
            {
                file = file + tag;
            }
            return folder + file + extension;
        }

        
    }
}
