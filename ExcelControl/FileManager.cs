using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelControl
{
    internal static class FileManager
    {
        public static void FileCopy(this FileInfo file, string destFolder)
        {
            file.CopyTo(Path.Combine(destFolder, file.Name), true);
        }
    }
}
