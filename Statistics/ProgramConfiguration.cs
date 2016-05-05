using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace Statistics
{
    public static class ProgramConfiguration
    {
        private static string docDownloadedFolder;
        private static string currentExcelFolder;
        private static string archivedExcelFolder;
        private static string archivedPdfFolder;
        private static string archivedCertificationFolder;

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
        #endregion
    }
}
