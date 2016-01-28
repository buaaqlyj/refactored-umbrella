using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics
{
    public class JobParameterStruct
    {
        private string startupPath;
        //历史记录处理
        private string inputFile;//文件完整路径
        private string inputFolder;//输入文件夹
        private string outputFolder;//输出文件夹
        //测试检定
        private string certDLFolder;//证书下载文件夹
        private string currentDataFolder;//当前检定数据存储文件夹
        private string historyDataFolder;//历史检定数据存储文件夹
        private string pdfDataFolder;//pdf检定数据存储文件夹
        private string certFolder;//证书存储文件夹
        private string tempFolder;//临时文件夹
        //选项
        private int dataPattern;//检定类型：kV,Dose,CT
        private int actionType;//操作类型：标准化，合并，生成证书
        private int fixType;//温压修正：不修正，修正
        private string macType;//仪器类型
        private string certTemp;//证书模板文件名
        private string dataTemp;//记录模板文件名
        private bool createNew;//对新记录是否建档

        public JobParameterStruct(string startupPath, string inputFi, string inputFo, string outputFo, string certDLFo, string currentDF, string historyDF, string pdfDF, string certFo, int pattern, int type, int fix, string macT, string certT, string dataT, bool createN)
        {
            this.startupPath = startupPath;

            inputFile = inputFi;
            inputFolder = inputFo;
            outputFolder = outputFo;

            certDLFolder = certDLFo;
            currentDataFolder = currentDF;
            historyDataFolder = historyDF;
            pdfDataFolder = pdfDF;
            certFolder = certFo;

            dataPattern = pattern;
            actionType = type;
            fixType = fix;
            macType = macT;
            certTemp = certT;
            dataTemp = dataT;
            createNew = createN;

            tempFolder = DataUtility.DataUtility.PathCombine(startupPath, @"Temp\");
            DataUtility.DataUtility.TryCreatFolder(tempFolder);
        }

        public string InputFile
        {
            set
            {
                inputFile = value;
            }
            get
            {
                return inputFile;
            }
        }

        public string InputFolder
        {
            set
            {
                inputFolder = value;
            }
            get
            {
                return inputFolder;
            }
        }

        public string OutputFolder
        {
            get
            {
                return outputFolder;
            }
            set
            {
                outputFolder = value;
            }
        }

        public string CertDLFolder
        {
            set
            {
                certDLFolder = value;
            }
            get
            {
                return certDLFolder;
            }
        }

        public string CurrentDataFolder
        {
            set
            {
                currentDataFolder = value;
            }
            get
            {
                return currentDataFolder;
            }
        }

        public string HistoryDataFolder
        {
            set
            {
                historyDataFolder = value;
            }
            get
            {
                return historyDataFolder;
            }
        }

        public string PDFDataFolder
        {
            set
            {
                pdfDataFolder = value;
            }
            get
            {
                return pdfDataFolder;
            }
        }

        public string CertFolder
        {
            set
            {
                certFolder = value;
            }
            get
            {
                return certFolder;
            }
        }

        public string TempFolder
        {
            get
            {
                return tempFolder;
            }
        }

        public int DataPattern
        {
            set
            {
                dataPattern = value;
            }
            get
            {
                return dataPattern;
            }
        }

        public int ActionType
        {
            get
            {
                return actionType;
            }
            set
            {
                actionType = value;
            }
        }

        public int FixType
        {
            get
            {
                return fixType;
            }
        }

        public string MacType
        {
            get
            {
                return macType;
            }
        }

        public string CertTemplateFilePath
        {
            get
            {
                return DataUtility.DataUtility.PathCombine(DataUtility.DataUtility.PathCombineClassified(startupPath + @"\试验证书模板", DataPattern), certTemp);
            }
        }

        public string DataTemplateFilePath
        {
            get
            {
                return DataUtility.DataUtility.PathCombine(DataUtility.DataUtility.PathCombineClassified(startupPath + @"\试验记录模板", DataPattern), dataTemp);
            }
        }

        public bool CreateNew
        {
            get
            {
                return createNew;
            }
        }

        public bool AutoFixType
        {
            get
            {
                if (fixType == 0) return false;
                else return true;
            }
        }

        public string AutoInputFolder
        {
            get
            {
                switch (actionType)
                {
                    case 0:
                        //下载好的证书的存放目录
                        return CertDLFolder;
                    case 1:
                        //生成空记录表的存放目录
                        return DataUtility.DataUtility.PathCombineClassified(CurrentDataFolder, DataPattern);
                    default:
                        return InputFolder;
                }
            }
        }

        public string AutoOutputFolder
        {
            get
            {
                switch (actionType)
                {
                    case 0:
                        //生成记录的存放位置
                        return DataUtility.DataUtility.PathCombineClassified(CurrentDataFolder, DataPattern);
                    case 1:
                        //历史存档记录的位置
                        return DataUtility.DataUtility.PathCombineClassified(HistoryDataFolder, DataPattern);
                    default:
                        return OutputFolder;
                }
            }
        }

        public string AutoExtension
        {
            get
            {
                switch (actionType)
                {
                    case 0:
                        return @"*.doc*";
                    default:
                        return @"*.xls*"; ;
                }
            }
        }
    }
}
