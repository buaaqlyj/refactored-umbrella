using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Util
{
    public static class PathExt
    {
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
    }
}
