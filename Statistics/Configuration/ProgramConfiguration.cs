using System;
using System.Collections.Generic;
using System.IO;

using Statistics.Instrument.Standard;
using Statistics.Instrument.Tested;
using Statistics.IO;
using Statistics.Log;

namespace Statistics.Configuration
{
    public static class ProgramConfiguration
    {
        private static string docDownloadedFolder;
        private static string currentExcelFolder;
        private static string archivedExcelFolder;
        private static string archivedPdfFolder;
        private static string archivedCertificationFolder;

        private static string dataTemplateFolder;
        private static string certTemplateFolder;

        private static string programFolder;

        private static Dictionary<string, StandardInstrument> standard = new Dictionary<string, StandardInstrument>();
        private static Dictionary<string, List<string>> standardUsage = new Dictionary<string, List<string>>();

        public static void Initial(string programFolder)
        {
            ProgramConfiguration.programFolder = programFolder;
            ProgramConfiguration.currentExcelFolder = programFolder + @"\当前实验记录";
            ProgramConfiguration.docDownloadedFolder = programFolder + @"\证书下载";
            ProgramConfiguration.archivedExcelFolder = programFolder + @"\历史数据记录";
            ProgramConfiguration.archivedCertificationFolder = programFolder + @"\证书记录";
            ProgramConfiguration.archivedPdfFolder = programFolder + @"\PDF数据记录";

            ProgramConfiguration.dataTemplateFolder = programFolder + @"\试验记录模板";
            ProgramConfiguration.certTemplateFolder = programFolder + @"\试验证书模板";

            DataUtility.DataUtility.TryCreatFolder(certTemplateFolder);
            DataUtility.DataUtility.TryCreatFolders(certTemplateFolder, new string[] { "CT", "KV", "剂量" });
            DataUtility.DataUtility.TryCreatFolder(dataTemplateFolder);
            DataUtility.DataUtility.TryCreatFolders(dataTemplateFolder, new string[] { "CT", "KV", "剂量" });
            DataUtility.DataUtility.TryCreatFolder(docDownloadedFolder);
            DataUtility.DataUtility.TryCreatFolder(archivedCertificationFolder);
            DataUtility.DataUtility.TryCreatFolder(currentExcelFolder);
            DataUtility.DataUtility.TryCreatFolders(currentExcelFolder, new string[] { "CT", "KV", "剂量" });
            DataUtility.DataUtility.TryCreatFolder(archivedPdfFolder);
            DataUtility.DataUtility.TryCreatFolder(archivedExcelFolder);
            DataUtility.DataUtility.TryCreatFolders(archivedExcelFolder, new string[] { "CT", "KV", "剂量" });

            TestedInstrument.InitialTypes(programFolder);
            DataUtility.DataUtility.TryCreatFolders(archivedExcelFolder + @"\CT", TestedInstrument.CTTypes);
            DataUtility.DataUtility.TryCreatFolders(archivedExcelFolder + @"\KV", TestedInstrument.KVTypes);
            DataUtility.DataUtility.TryCreatFolders(archivedExcelFolder + @"\剂量", TestedInstrument.DoseTypes);

            ExpiredDataValidate();
        }

        private static void ExpiredDataValidate()
        {
            FileInfo fi = new FileInfo(programFolder + @"\ExpiredDate.ini");
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
                LogHelper.AddDataError("找不到ExpiredDate.ini文件", true);
            }
        }

        #region Property
        public static string DocDownloadedFolder
        {
            get
            {
                return docDownloadedFolder;
            }
            set
            {
                docDownloadedFolder = value;
            }
        }

        public static string CurrentExcelFolder
        {
            get
            {
                return currentExcelFolder;
            }
            set
            {
                currentExcelFolder = value;
            }
        }

        public static string ArchivedExcelFolder
        {
            get
            {
                return archivedExcelFolder;
            }
            set
            {
                archivedExcelFolder = value;
            }
        }

        public static string ArchivedPdfFolder
        {
            get
            {
                return archivedPdfFolder;
            }
            set
            {
                archivedPdfFolder = value;
            }
        }

        public static string ArchivedCertificationFolder
        {
            get
            {
                return archivedCertificationFolder;
            }
            set
            {
                archivedCertificationFolder = value;
            }
        }

        public static string DataTemplateFolder
        {
            get
            {
                return dataTemplateFolder;
            }
            set
            {
                dataTemplateFolder = value;
            }
        }

        public static string CertTemplateFolder
        {
            get
            {
                return certTemplateFolder;
            }
            set
            {
                certTemplateFolder = value;
            }
        }

        public static string ProgramFolder
        {
            get
            {
                return programFolder;
            }
            set
            {
                programFolder = value;
            }
        }
        #endregion
    }
}
