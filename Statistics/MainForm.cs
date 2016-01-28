using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading;
using System.Reflection;

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
    /// <summary>
    /// 标准化：
    /// 1、打开标准表，重新导入每年的数据。成功导入的删除原文件
    /// 2、名称要求：年份+单位名称+仪器名称+仪器编号+“探测器”+探测器编号
    /// 
    /// 合并：
    /// 1、比较：单位名称+仪器编号+探测器编号（不区分大小写和符号空格）。成功导入的删除原文件
    /// 报错逻辑①后两者相同，单位名称不同，报错人为错误
    /// 报错逻辑②一三相同，仪器编号不同，报错人为错误
    /// 2、合并逻辑：复制第一年（08年），复制一页标准sheet，逐年合并后年数据
    /// 3、合并完，正常下删除原数据，只保留新的。
    /// 4、合并完，比较送检单位，仪器名称，仪器编号，电离室编号，不同就报错
    /// 5、检查证书编号，没有报错。
    /// 
    /// 问题：
    /// 1、有问题的不操作，没问题的删掉
    /// 2、022没报错
    /// 3、关闭excel
    /// 
    /// 2015/1/20 问题
    /// 1、超差需要保存在日志
    /// 2、校验加入重新统计功能
    /// 3、文件名相重，提示是否替换
    /// 4
    /// 
    /// 2015/5/7 project
    /// 1.        CT剂量的存档合并功能，着重研究统计页面不同规范与对应列的关系和统计长期稳定性时遇到无数据的安全处理办法；
    /// 2.【完成】CopyData中完善签名拷贝和生成的代码部分；生成excel是记录者签名，合并校验生成证书是校对者签名，pdf里有两个签名
    /// 3.【完成】完善文件夹部分代码，添加仪器类型与仪器文件夹的对应关系；
    /// 4.        CT剂量的正常检定功能。
    /// 5.        生成证书拷贝备注
    /// 
    /// 2015/11/27 新计划
    /// 重构旧代码
    /// </summary>
    public partial class MainForm : Form
    {
        bool checkClear = false;
        //StreamWriter logFile = null;
        bool checkSuperDog = false;
        public bool isWorking, isStopping;
        StreamWriter logFile = null;
        Action<JobParameterStruct> mi;
        IAsyncResult ar;
        delegate void doOneDelegate(string filePath, JobParameterStruct paraStruct, ref List<string> problemFilesList);
        doOneDelegate dod;
        string currentFile = "";
        int exceptionNum = 0;
        int dataerrorNum = 0;
        FileInfo[] existFile = null;
        SuperDog.Person person = null;
        Dictionary<string, StandardInstrument> standard = new Dictionary<string, StandardInstrument>();
        Dictionary<string, List<string>> standardUsage = new Dictionary<string, List<string>>();
        
        public MainForm()
        {
            InitializeComponent();
        }
        
        #region Log

        public void Log_Write(string log)
        {
            if (logFile != null)
            {
                logFile.WriteLine(log);
                logFile.Flush();
            }
        }

        public void AddLog(string pre, string ex, bool sw)
        {
            string temp = @"【" + pre + @"】" + ex;
            if (sw)
            {
                Log_Write(temp);
            }
            TextBox_Write(temp + Environment.NewLine);
        }

        public void AddLog(string ex, bool sw)
        {
            if (sw)
            {
                Log_Write(ex);
            }
            TextBox_Write(ex + Environment.NewLine);
        }

        #endregion

        #region Invoke&Delegate&Event
        
        public void AddDataError(string ex, bool log)
        {
            dataerrorNum++;
            if (exceptionNum + dataerrorNum == 1)
            {
                AddLog(@"信息18", "原文件名：" + currentFile, true);
            }
            AddLog(@"错误19", "  第" + dataerrorNum + "个数据错误：" + ex, log);
        }

        public void AddException(string ex, bool log)
        {
            exceptionNum++;
            if (exceptionNum + dataerrorNum == 1)
            {
                AddLog(@"信息20", "原文件名：" + currentFile, true);
            }
            AddLog(@"错误01", "  第" + exceptionNum + "个格式错误：" + ex, log);
        }

        delegate void SignitureWriteInvoke(Bitmap sig);
        public void Signitue_Write(Bitmap sig)
        {
            if (pictureBox1.InvokeRequired)
            {
                SignitureWriteInvoke lwi = new SignitureWriteInvoke(Signitue_Write);
                this.BeginInvoke(lwi, new object[] { sig });
            }
            else
            {
                pictureBox1.Image = sig;
                //pictureBox1.Refresh();
            }
        }

        delegate void TextBoxWriteInvoke(string str);
        public void TextBox_Write(string log)
        {
            if (textBox2.InvokeRequired)
            {
                TextBoxWriteInvoke tbwi = new TextBoxWriteInvoke(TextBox_Write);
                this.BeginInvoke(tbwi, new object[] { log });
            }
            else
            {
                textBox2.AppendText(log);
                textBox2.ScrollToCaret();
            }
        }

        delegate void ProgressBarSetValueInvoke(double value);
        delegate void LabelDisplayProgressInvoke(double value);
        public void ProgressBar_SetValue(double value)
        {
            if (progressBar1.InvokeRequired)
            {
                ProgressBarSetValueInvoke pbsi = new ProgressBarSetValueInvoke(ProgressBar_SetValue);
                this.BeginInvoke(pbsi, new object[] { value });
            }
            else
            {
                progressBar1.Value = (int)Math.Ceiling(value);
            }
        }
        public void Label_DisplayProgress(double value)
        {
            if (label5.InvokeRequired)
            {
                LabelDisplayProgressInvoke ldpi = new LabelDisplayProgressInvoke(Label_DisplayProgress);
                this.BeginInvoke(ldpi, new object[] { value });
            }
            else
            {
                label5.Text = value.ToString(@"0.00") + @" %";
            }
        }
        public void UpdateProgress(double value)
        {
            ProgressBar_SetValue(Math.Max((int)0, (int)Math.Min((int)100, (int)Math.Ceiling(value))));
            Label_DisplayProgress(value);
        }

        delegate void ButtonSetTextInvoke(string text);
        public void Button_SetText(string text)
        {
            if (button1.InvokeRequired)
            {
                ButtonSetTextInvoke bsti = new ButtonSetTextInvoke(Button_SetText);
                this.BeginInvoke(bsti, new object[] { text });
            }
            else
            {
                button1.Text = text;
            }
        }

        delegate void ToolStripStatusLabelSetTextInvoke(string text);
        public void ToolStripStatusLabel_SetText(string text)
        {
            if (statusStrip1.InvokeRequired)
            {
                ToolStripStatusLabelSetTextInvoke tsslsti = new ToolStripStatusLabelSetTextInvoke(ToolStripStatusLabel_SetText);
                this.BeginInvoke(tsslsti, new object[] { text });
            }
            else
            {
                toolStripStatusLabel1.Text = text;
            }
        }

        delegate void ToolStripStatusProgressBarSetValueInvoke(double value);
        public void ToolStripStatusProgressBar_SetValue(double value)
        {
            if (statusStrip1.InvokeRequired)
            {
                ToolStripStatusProgressBarSetValueInvoke tsspbsvi = new ToolStripStatusProgressBarSetValueInvoke(ToolStripStatusProgressBar_SetValue);
                this.BeginInvoke(tsspbsvi, new object[] { value });
            }
            else
            {
                toolStripProgressBar1.Value = Math.Max((int)0, (int)Math.Min((int)100, (int)Math.Ceiling(value)));
                toolStripStatusLabel2.Text = value.ToString(@"0.00") + @" %";
            }
        }
        #endregion

        #region Control

        private void timer1_Tick(object sender, EventArgs e)
        {
            SuperDog.SuperDogSeries myDog = new SuperDog.SuperDogSeries();
            if (checkSuperDog && !myDog.DogFlag)
            {
                timer1.Stop();
                MessageBox.Show("Error:     " + myDog.Status.ToString() + "\nSuperDog disabled!");
                this.Close();
                return;
            }
        }
        /// <summary>
        /// “~$湖北省中职检测研究院-2014-2531.xlsx”
        /// 
        /// 
        /// 3种状态：
        ///             isWorking isStopping button1.Text
        /// 1停止未工作：false     false     统计
        /// 2正在工作中：true      false     停止
        /// 3等待退出中：true      true      等待当前工作停止...
        /// 
        /// 1-2点击开始：isWorking = true;
        /// 2-3点击停止：isStopping = false;
        /// 3-1线程退出：isWorking = false; isStopping = false;
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (isWorking)
            {
                if (isStopping)
                {
                    //3-等待退出中：true      true      等待当前工作停止...
                    //isWorking = false;
                } 
                else
                {
                    // 2-正在工作中：true      false     停止
                    //  |
                    //  V
                    // 3-等待退出中：true      true      等待当前工作停止...
                    isStopping = true;
                    button1.Text = "等待当前工作停止...";
                }
            }
            else
            {
                if (!isStopping)
                {
                    //1-停止未工作：false     false     统计
                    // |
                    // V
                    //2-正在工作中：true      false     停止
                    isWorking = true;
                    button1.Text = "停止";

                    ar = mi.BeginInvoke(new JobParameterStruct(Application.StartupPath, textBox1.Text, textBox7.Text, textBox6.Text, ProgramConfiguration.DocDownloadedFolder, ProgramConfiguration.CurrentExcelFolder, ProgramConfiguration.ArchivedExcelFolder, ProgramConfiguration.ArchivedPdfFolder, ProgramConfiguration.ArchivedCertificationFolder, comboBox1.SelectedIndex, comboBox2.SelectedIndex, comboBox6.SelectedIndex, comboBox3.Text, comboBox5.Text, comboBox4.Text, checkBox1.Checked), null, null);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Excel文档(*.xlsx,*.xls)|*.xls*|所有文件(*.*)|*.*";
            this.openFileDialog1.FileName = textBox1.Text;

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = this.openFileDialog1.FileName;
                textBox1.Text = FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "请选择文件路径";
            folderBrowserDialog1.SelectedPath = textBox7.Text;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string foldPath = folderBrowserDialog1.SelectedPath;
                textBox7.Text = foldPath;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "请选择文件路径";
            folderBrowserDialog1.SelectedPath = textBox6.Text;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string foldPath = folderBrowserDialog1.SelectedPath;
                textBox6.Text = foldPath;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(ProgramConfiguration.DocDownloadedFolder);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(ProgramConfiguration.CurrentExcelFolder);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(ProgramConfiguration.ArchivedExcelFolder);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(ProgramConfiguration.ArchivedPdfFolder);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(ProgramConfiguration.ArchivedCertificationFolder);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox7.Text = "";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBox6.Text);
        }
        //数据类型:CT/KV/剂量
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string currentInstrument = "";
            if (comboBox3.SelectedItem != null)
            {
                currentInstrument = comboBox3.SelectedItem.ToString().ToLower();
            }
            switch (comboBox1.SelectedIndex)
            {
                case 0://剂量
                    comboBox3.Items.Clear();
                    foreach (string item in TestedInstrument.DoseTypes)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case 1://CT
                    comboBox3.Items.Clear();
                    foreach (string item in TestedInstrument.CTTypes)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case 2://KV
                    comboBox3.Items.Clear();
                    foreach (string item in TestedInstrument.KVTypes)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
            }
            comboBox3.SelectedIndex = 0;
            for (int i = 0; i < comboBox3.Items.Count; i++)
            {
                if (comboBox3.Items[i].ToString().ToLower() == currentInstrument)
                {
                    comboBox3.SelectedIndex = i;
                    break;
                }
            }
            updateSelectionAndLabel();
        }
        //操作类型:生成记录/生成证书/标准化/存档合并/校验
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateSelectionAndLabel();
        }
        //温压修正:半导体/电离室
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateSelectionAndLabel();
        }

        //仪器类型:
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateSelectionAndLabel();
        }
        //记录模板:CT/KV/剂量
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //证书模板:CT/KV/剂量+半导体/电离室
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (!isWorking)
            {
                existFile = (new DirectoryInfo(DataUtility.DataUtility.PathCombineClassified(ProgramConfiguration.ArchivedExcelFolder, comboBox1.SelectedIndex))).GetFiles(@"*.xls*", SearchOption.AllDirectories);
            }
        }

        private void updateSelectionAndLabel()
        {
            string dpath = DataUtility.DataUtility.PathCombineClassified(System.Windows.Forms.Application.StartupPath + @"\试验记录模板\", comboBox1.SelectedIndex);
            string cpath = DataUtility.DataUtility.PathCombineClassified(System.Windows.Forms.Application.StartupPath + @"\试验证书模板\", comboBox1.SelectedIndex);
            string keyword = "";
            string dataType = "";
            //数据类型
            switch (comboBox1.SelectedIndex)
            {
                case 0://剂量
                    dataType = "Dose";
                    break;
                case 1://CT
                    dataType = "CT";
                    break;
                case 2://KV
                    dataType = "KV";
                    switch (comboBox3.Text)
                    {
                        case "4000M+":
                        case "TNT12000":
                        case "TNT12000D":
                            keyword = comboBox3.Text;
                            dpath = Path.Combine(dpath, @"FLUKE");
                            break;
                    }
                    break;
                default:
                    MessageBox.Show(@"没有选择类型：剂量/CT/KV");
                    break;
            }
            //设置记录模板
            DirectoryInfo mydir = new DirectoryInfo(dpath);
            FileInfo[] fis = mydir.GetFiles();
            comboBox4.Items.Clear();
            if (fis.Length == 0)
            {
                MessageBox.Show("记录模板文件夹为空，请检查！");
            }
            else
            {
                foreach (FileInfo fi in fis)
                {
                    if (fi.Name.Contains(keyword) && !fi.Name.Contains(@"~$"))
                    {
                        comboBox4.Items.Add(fi.Name);
                    }
                }
                if (comboBox4.Items.Count > 0)
                {
                    comboBox4.SelectedIndex = 0;
                }
                else
                {
                    MessageBox.Show("记录模板文件夹为空，请检查！");
                }
            }
            //设置证书模板
            mydir = new DirectoryInfo(cpath);
            fis = mydir.GetFiles();
            comboBox5.Items.Clear();
            if (fis.Length == 0)
            {
                MessageBox.Show("证书模板文件夹为空，请检查！");
            }
            else
            {
                foreach (FileInfo fi in fis)
                {
                    if (fi.Name.Contains(keyword) && !fi.Name.Contains(@"~$"))
                    {
                        comboBox5.Items.Add(fi.Name);
                    }
                }
                if (comboBox5.Items.Count > 0)
                {
                    comboBox5.SelectedIndex = 0;
                }
                else
                {
                    MessageBox.Show("证书模板文件夹为空，请检查！");
                }
            }
            //更新颜色显示
            if (comboBox2.SelectedIndex == 0)
            {
                //生成记录
                label7.ForeColor = Color.Red;
                if (comboBox1.SelectedIndex == 2)
                {
                    label15.ForeColor = Color.Black;
                }
                else
                {
                    label15.ForeColor = Color.Red;
                }
                label9.ForeColor = Color.Red;
                label10.ForeColor = Color.Black;
                checkBox1.ForeColor = Color.Black;
                dod = new doOneDelegate(GeneratingForm);
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                //生成证书
                label7.ForeColor = Color.Black;
                label15.ForeColor = Color.Black;
                label9.ForeColor = Color.Black;
                label10.ForeColor = Color.Red;
                checkBox1.ForeColor = Color.Red;
                dod = new doOneDelegate(GeneratingCertificateOne);
            }
            else if (comboBox2.SelectedIndex == 2)
            {
                //标准化
                label7.ForeColor = Color.Black;
                label15.ForeColor = Color.Black;
                label9.ForeColor = Color.Red;
                label10.ForeColor = Color.Black;
                checkBox1.ForeColor = Color.Black;
                dod = new doOneDelegate(StandardizeOne);
            }
            else if (comboBox2.SelectedIndex == 3)
            {
                //存档合并
                label7.ForeColor = Color.Black;
                label15.ForeColor = Color.Black;
                label9.ForeColor = Color.Black;
                label10.ForeColor = Color.Black;
                checkBox1.ForeColor = Color.Black;
                dod = new doOneDelegate(ArchievingMergeOne);
            }
            else if (comboBox2.SelectedIndex == 4)
            {
                //校验
                label7.ForeColor = Color.Black;
                label15.ForeColor = Color.Black;
                label9.ForeColor = Color.Black;
                label10.ForeColor = Color.Black;
                checkBox1.ForeColor = Color.Black;
                dod = new doOneDelegate(VerificateOne);
            }

            //检查所涉及的标准仪器是否超期
            if (comboBox1.SelectedIndex > -1)
            {
                List<ListViewItem> lvItem = new List<ListViewItem>();
                ListViewItem lvi = null;
                Color textColor = Color.Black;
                int abnormalItemCount = 0;
                foreach (string item in standardUsage[dataType])
                {
                    if (standard[item].State == "正常")
                    {
                        textColor = Color.Green;
                    }
                    else if (standard[item].State == "即将过期")
                    {
                        textColor = Color.DarkOrange;
                        abnormalItemCount++;
                    }
                    else
                    {
                        textColor = Color.Red;
                        abnormalItemCount++;
                    }
                    lvi = new ListViewItem(new string[] { standard[item].Name, standard[item].Date, standard[item].State }, -1, textColor, System.Drawing.Color.Empty, null);
                    lvItem.Add(lvi);
                }
                listView1.Items.Clear();
                foreach (ListViewItem item in lvItem)
                {
                    listView1.Items.Add(item);
                }
                if (abnormalItemCount > 0)
                {
                    label19.Text = "有标准仪器已超期或即将超期！";
                    label19.ForeColor = Color.DarkOrange;
                }
                else
                {
                    label19.Text = "标准仪器状态正常！";
                    label19.ForeColor = Color.Green;
                }
            }
        }

        #endregion

        #region Form

        private void MainForm_Load(object sender, EventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath + @"\日志");
            if (!di.Exists)
            {
                di.Create();
            }
            logFile = new StreamWriter(Application.StartupPath + @"\日志\" + DateTime.Now.ToString(@"yyyyMMdd-HH-mm-ss") + @".txt");
            //LogHelper.Init();

            ExcelUtility.lfwi += new ExcelUtility.LogFileWriteInvoke(this.Log_Write);
            ExcelUtility.tbwi += new ExcelUtility.TextBoxWriteInvoke(this.TextBox_Write);
            ExcelUtility.aed += new ExcelUtility.AddExceptionDelegate(this.AddException);
            ExcelUtility.aded += new ExcelUtility.AddDataErrorDelegate(this.AddDataError);

            WordUtility.lfwi += new WordUtility.LogFileWriteInvoke(this.Log_Write);
            WordUtility.tbwi += new WordUtility.TextBoxWriteInvoke(this.TextBox_Write);
            WordUtility.aed += new WordUtility.AddExceptionDelegate(this.AddException);
            WordUtility.aded += new WordUtility.AddDataErrorDelegate(this.AddDataError);

            DataUtility.DataUtility.lfwi += new DataUtility.DataUtility.LogFileWriteInvoke(this.Log_Write);
            DataUtility.DataUtility.tbwi += new DataUtility.DataUtility.TextBoxWriteInvoke(this.TextBox_Write);
            DataUtility.DataUtility.aed += new DataUtility.DataUtility.AddExceptionDelegate(this.AddException);
            DataUtility.DataUtility.aded += new DataUtility.DataUtility.AddDataErrorDelegate(this.AddDataError);

            Log.LogHelper.AddDataErrorEvent += new Log.LogHelper.AddDataErrorHandler(AddDataError);
            Log.LogHelper.AddExceptionEvent += new Log.LogHelper.AddExceptionHandler(AddException);
            Log.LogHelper.AddLogEvent += new Log.LogHelper.AddLogHandler(AddLog);

            //DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\日志");
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\试验证书模板");
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\试验证书模板", new string[] { "CT", "剂量", "KV" });
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\试验记录模板");
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\试验记录模板", new string[] { "CT", "剂量", "KV" });
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\证书下载");
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\证书记录");
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\当前实验记录");
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\当前实验记录", new string[] { "CT", "剂量", "KV" });
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\PDF数据记录");
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\历史数据记录");
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\历史数据记录", new string[] { "CT", "剂量", "KV" });
            DataUtility.DataUtility.TryCreatFolder(Application.StartupPath + @"\输出文件夹");

            TestedInstrument.InitialTypes(Application.StartupPath);
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\历史数据记录\CT", TestedInstrument.CTTypes);
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\历史数据记录\KV", TestedInstrument.KVTypes);
            DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\历史数据记录\剂量", TestedInstrument.DoseTypes);
            //DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\当前实验记录\剂量", existType);
            //DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\当前实验记录\CT", existType);
            //DataUtility.DataUtility.TryCreatFolders(Application.StartupPath + @"\当前实验记录\KV", existType);
            //comboBox2.SelectedIndex = 0;

            SuperDog.SuperDogSeries myDog = new SuperDog.SuperDogSeries();
            myDog.RunDogTesting();
            if (myDog.DogFlag)
            {
                Signitue_Write(myDog.Signiture);
                person = myDog.Person;
            }
            else if (checkSuperDog)
            {
                MessageBox.Show("Error:     " + myDog.Status.ToString() + "\nSuperDog disabled!");
                this.Close();
                return;
            }
            else
            {
                Signitue_Write(SuperDog.SuperDogAlpha.getInstance().Signiture);
                person = SuperDog.SuperDogAlpha.getInstance().Person;
            }

            FileInfo fi = new FileInfo(Application.StartupPath + @"\ExpiredDate.ini");
            if (fi.Exists)
            {
                string[] sections = INI.INIGetAllSectionNames(fi.FullName);
                string[] keys = null;
                List<string> instrument = new List<string>();
                int tempInt;
                foreach (string item in sections)
                {
                    if (Int32.TryParse(item, out tempInt))
                    {
                        standard.Add(item, new StandardInstrument(INI.INIGetStringValue(fi.FullName, item, "Name", null), INI.INIGetStringValue(fi.FullName, item, "Date", null)));
                    }
                    else
                    {
                        keys = INI.INIGetAllItemKeys(fi.FullName, item);
                        instrument = new List<string>();
                        foreach (string item1 in keys)
                        {
                            instrument.Add(INI.INIGetStringValue(fi.FullName, item, item1, null));
                        }
                        if (standardUsage.ContainsKey(item))
                        {
                            standardUsage.Remove(item);
                        }
                        standardUsage.Add(item, instrument);
                    }
                }
            }
            else
            {
                AddDataError("找不到ExpiredDate.ini文件", true);
            }

            isWorking = false;
            isStopping = false;
            mi = DoWork;
            dod = new doOneDelegate(StandardizeOne);

            textBox6.Text = Application.StartupPath + @"\输出文件夹";
            ProgramConfiguration.CurrentExcelFolder = Application.StartupPath + @"\当前实验记录"; //textBox3
            ProgramConfiguration.DocDownloadedFolder = Application.StartupPath + @"\证书下载"; //textBox4
            ProgramConfiguration.ArchivedExcelFolder = Application.StartupPath + @"\历史数据记录"; //textBox5
            ProgramConfiguration.ArchivedCertificationFolder = Application.StartupPath + @"\证书记录"; //textBox8
            ProgramConfiguration.ArchivedPdfFolder = Application.StartupPath + @"\PDF数据记录"; //textBox9

            comboBox1.SelectedIndex = 0;
            
            comboBox3.SelectedIndex = 0;

            timer1.Start();
        }

        private void MainForm_FormClosing(object sender, EventArgs e)
        {
            try
            {
                logFile.Close();
                DataUtility.DataUtility.TryDeleteFilesInFolders(Application.StartupPath + @"\试验记录模板", @"~$");
                DataUtility.DataUtility.TryDeleteFilesInFolders(Application.StartupPath + @"\试验证书模板", @"~$");
            }
            catch
            {

            }
        }

        #endregion

        #region Operation
        
        public void DoWork(JobParameterStruct pS)
        {
            string inputFi = pS.InputFile;
            string inputFo = pS.AutoInputFolder;
            string outputFo = pS.AutoOutputFolder;
            string inputEx = pS.AutoExtension;
            int cbsi = pS.DataPattern;

            if (pS.ActionType < 0 || pS.ActionType > 4)
            {
                if (MessageBox.Show(@"没有选择任何操作类型，请点击确定后返回重试", "提示", MessageBoxButtons.OK) == DialogResult.OK) { }
            }
            else if (pS.ActionType < 1 && (pS.FixType < 0 || pS.FixType > 1))
            {
                if (MessageBox.Show(@"没有选择温度气压修正选项，请点击确定后返回重试", "提示", MessageBoxButtons.OK) == DialogResult.OK) { }
            }
            else if (MessageBox.Show(@"处理成功将会删除原纪录，是否要继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                List<string> probfile = new List<string>();
                UpdateProgress(0);

                if (inputFi != "")
                {
                    if (!inputFi.Contains(@"~$"))
                    {
                        currentFile = inputFi;
                        ToolStripStatusLabel_SetText(new FileInfo(inputFi).Name);
                        try
                        {
                            dataerrorNum = 0;
                            exceptionNum = 0;
                            dod(inputFi, pS, ref probfile);
                        }
                        catch (Exception ex)
                        {
                            probfile.Add(inputFi);
                            if (exceptionNum > 0)
                            {
                                exceptionNum = 0;
                            }
                            else
                            {
                                AddLog(@"***************************************************************", true);
                            }
                            AddLog(@"* 文件名称：" + inputFi, true);
                            AddLog(@"* 异常消息：" + ex.Message, true);
                            AddLog(@"* 异常方法：" + ex.TargetSite, true);
                            AddLog(@"***************************************************************", true);
                        }
                        if (dataerrorNum > 0)
                        {
                            dataerrorNum = 0;
                            AddLog(@"***************************************************************", true);
                        }
                        ToolStripStatusLabel_SetText("");
                    }

                    UpdateProgress(100);

                    if (probfile.Count > 0)
                    {
                        AddLog(@"结果02", "共处理文件数：1", true);
                        AddLog(@"结果03", "异常文件总数：" + probfile.Count, true);
                        foreach (string item in probfile)
                        {
                            AddLog(@"结果04", item, true);
                        }
                    }
                    else
                    {
                        AddLog(@"结果05", "没有发现任何异常，共处理文件数：1", true);
                    }
                    MessageBox.Show("共处理文件" + "1" + "个。" + Environment.NewLine + "其中成功" + (1-probfile.Count).ToString() + "个，失败" + probfile.Count + "个。", "处理完成", MessageBoxButtons.OK);
                }
                else if (inputFo != "")
                {
                    DirectoryInfo di = new DirectoryInfo(inputFo);
                    FileInfo[] fis = di.GetFiles(inputEx, SearchOption.AllDirectories);
                    probfile.Clear();
                    int doneNumber = 0;
                    if (fis.Length > 0)
                    {
                        double Step = 100.0 / (double)fis.Length;
                        double values = 0;
                        foreach (FileInfo item in fis)
                        {
                            if (!item.Name.StartsWith(@"~$"))
                            {
                                currentFile = item.FullName;
                                ToolStripStatusLabel_SetText(new FileInfo(item.FullName).Name);
                                try
                                {
                                    dataerrorNum = 0;
                                    exceptionNum = 0;
                                    dod(item.FullName, pS, ref probfile);
                                }
                                catch (Exception ex)
                                {
                                    probfile.Add(item.FullName);
                                    if (exceptionNum > 0)
                                    {
                                        exceptionNum = 0;
                                    }
                                    else
                                    {
                                        AddLog(@"***************************************************************", true);
                                    }
                                    AddLog(@"* 文件名称：" + item.FullName, true);
                                    AddLog(@"* 异常消息：" + ex.Message, true);
                                    AddLog(@"* 异常方法：" + ex.TargetSite, true);
                                    AddLog(@"***************************************************************", true);
                                }
                                if (dataerrorNum > 0)
                                {
                                    dataerrorNum = 0;
                                    AddLog(@"***************************************************************", true);
                                }
                                ToolStripStatusLabel_SetText("");
                            }
                            values += Step;
                            UpdateProgress(values);
                            doneNumber++;
                            if (isStopping)
                            {
                                break;
                            }
                        }
                        if (probfile.Count > 0)
                        {
                            AddLog(@"结果06", "异常文件总数：" + probfile.Count, true);
                            foreach (string item in probfile)
                            {
                                AddLog(@"结果07", item, true);
                            }
                        }
                        else
                        {
                            AddLog(@"结果08", "没有发现任何异常，共处理文件数：" + doneNumber, true);
                        }
                        MessageBox.Show("共处理文件" + fis.Length + "个。" + Environment.NewLine + "其中成功" + (fis.Length - probfile.Count).ToString() + "个，失败" + probfile.Count + "个。", "处理完成", MessageBoxButtons.OK);
                    }
                    else
                    {
                        MessageBox.Show(@"输入文件夹没有找到待处理的文件");
                    }
                }
                else
                {
                    MessageBox.Show(@"未指定输入文件");
                }
            }
            isWorking = false;
            isStopping = false;
            Button_SetText(@"开始");
            AddLog(Environment.NewLine + Environment.NewLine, true);
        }

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
            JobMethods.GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
            //判断目标行数
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
                existFile = JobMethods.GetFilesFromType(output, _sr.GetText(_sr.ExcelWorkbook, stateIndex, "F5", out checkClear), ext, out checkClear);
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
                            newName = DataUtility.DataUtility.PathRightFileName(output, tempName, fi.Extension, "_new");
                            _sr.ExcelWorkbook.SaveAs(newName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            temp_fi = new FileInfo(newName);
                            JobMethods.Statistic(_sr, pattern, Perfect, strCompany, strType, tempName, certId);
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

                    JobMethods.CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out checkClear);
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
                    canGeCe = JobMethods.Statistic(_eu, pattern, Perfect, strCompany, strType, tempName, certId);
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
                        JobMethods.GenerateCert(_sr, stateIndex, pS.DataPattern, pS.CertTemplateFilePath, pS.CertFolder, pS.PDFDataFolder, pS.TempFolder, shouldFix, out success);
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

            JobMethods.TypeStandardize(_sr, stateIndex);

            certId = exSheets[stateIndex];
            year = certId.Substring(0, 4);

            tempName = year + "_" + _sr.GenerateFileName(_sr.ExcelWorkbook, stateIndex, out temp_strCompany, out temp_strType, out temp_strMacSerial, out temp_strSensorSerial, out Perfect);
            if (temp_strCompany == "") AddException(@"送校单位信息未提取到", true);
            if (temp_strType == "") AddException(@"仪器型号信息未提取到", true);
            if ((temp_strSensorSerial.Length + temp_strMacSerial.Length) == 0) AddException(@"主机编号和探头编号信息未提取到", true);
            if (temp_strMacSerial == "" || temp_strMacSerial == "/") AddException(@"主机编号信息未提取到", true);

            JobMethods.GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
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
                JobMethods.CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out newSheetIndex, out checkClear);
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
            JobMethods.GetFixState(_sr, stateIndex, pattern, out needFix, out shouldFix);
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
                temp_fi = JobMethods.SearchForFile(filePath, strType, strMacSerial, strSensorSerial, existFile, out isArchieved);
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
                        JobMethods.Statistic(_sr, pattern, Perfect, strCompany, strType, tempName, certId);
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

                    JobMethods.CopyData(_sr, stateIndex, _eu, pattern, certId, needFix, shouldFix, startDestiRowIndex, out checkClear);
                    if (checkClear)
                    {
                        _eu.ExcelWorkbook.Save();
                        _eu.ExcelWorkbook.Saved = true;
                        JobMethods.Statistic(_eu, pattern, Perfect, strCompany, strType, tempName, certId);
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