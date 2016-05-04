using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Util
{
    public static class PathExt
    {
        #region Combine
        //Override default Combine method because it returns short path when the path is under applications' folder.
        public static string PathCombine(string folder1, string folder2)
        {
            string combinedName = folder1 + @"\" + folder2;
            while (combinedName.Contains(@"\\"))
            {
                combinedName = combinedName.Replace(@"\\", @"\");
            }
            return combinedName;
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
                    //检定
                    keyword = "检定";
                    break;
                case 1:
                    //校准
                    keyword = "校准";
                    break;
            }
            return PathCombine(folder, keyword);
        }
        #endregion

        #region Change
        public static string PathChangeDirectory(string path, string directory)
        {
            return PathCombine(directory, PathGetFileName(path));
        }
        #endregion

        #region Base Member
        public static string PathChangeExtension(string path, string extension)
        {
            return Path.ChangeExtension(path, extension);
        }
        
        public static string PathGetFileName(string file)
        {
            return Path.GetFileName(file);
        }
        #endregion
    }
}
