using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Helpers
{
    public static class FileOperations
    {
        #region Move File
        public static void MoveFile(DirectoryInfo SourceDirectory, string SourceFile, DirectoryInfo DestinationDirectory, string DestinationFile)
        {
            string SourcePath = Path.Combine(SourceDirectory.FullName, SourceFile);
            string DestinationPath = Path.Combine(DestinationDirectory.FullName, DestinationFile);

            if (File.Exists(SourcePath))
            {
                File.Copy(SourcePath, DestinationPath,true);
                File.Delete(SourcePath);
            }
        }
        #endregion
    }
}