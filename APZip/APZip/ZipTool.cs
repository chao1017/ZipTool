using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Ionic.Zip;

namespace APZip
{
    public class ZipTool
    {
        #region Zip Files
        /// <summary>
        /// 壓縮檔案
        /// </summary>
        /// <param name="iToZipFiles">要壓縮的檔案</param>
        /// <param name="iPassword">壓縮密碼，null or empty 則表示沒密碼</param>
        /// <param name="iZipFile">壓縮後的檔名</param>
        /// <returns></returns>
        public bool ZipFiles(string[] iToZipFiles, string iZipFile, string iPassword)
        {
            bool Result = false;

            try
            {
                using (ZipFile zip = new ZipFile())
                {
                    if (!string.IsNullOrEmpty(iPassword))
                        zip.Password = iPassword;

                    //zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed;
                    foreach (string item in iToZipFiles)
                    {
                        zip.AddFile(item);
                    }

                    if (Path.GetExtension(iZipFile) == "")   //表示沒有輸入附檔名，則加上 .zip
                        iZipFile += ".zip";

                    zip.Save(iZipFile);
                    Result = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Result;
        }
        #endregion

        #region Zip Folders
        /// <summary>
        /// 壓縮目錄
        /// </summary>
        /// <param name="iToZipFolders">來源的目錄</param>
        /// <param name="iPassword">壓縮密碼，null or empty 則表示沒密碼</param>
        /// <param name="iZipFile">壓縮後的檔名</param>
        /// <returns></returns>
        public bool ZipFolders(string[] iToZipFolders, string iZipFile, string iPassword)
        {
            bool Result = false;

            try
            {
                using (ZipFile zip = new ZipFile())
                {
                    if (!string.IsNullOrEmpty(iPassword))
                        zip.Password = iPassword;

                    //zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestSpeed;
                    foreach (string item in iToZipFolders)
                    {
                        zip.AddDirectory(item);
                    }

                    if (string.IsNullOrEmpty(Path.GetExtension(iZipFile)))   //表示沒有輸入附檔名，則加上 .zip
                        iZipFile += ".zip";

                    zip.Save(iZipFile);
                    Result = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Result;
        }
        #endregion

        #region Unzip Files
        /// <summary>
        /// 檔案解壓縮
        /// </summary>
        /// <param name="iZipFile">壓縮檔</param>
        /// <param name="iUnZipPath">解壓縮後的路徑，如果不存在會自動建立目錄</param>
        /// <param name="iCreateRelativePath">是否要依據來源的相對路徑建立在目標目錄下</param>
        /// <param name="iOverWriteExistFile">如果檔案存在，是否要覆蓋</param>
        /// <param name="iPassword">解壓縮密碼，Null or Empty 表示不需要密碼</param>
        /// <returns></returns>
        //public bool UnZipFile(string iZipFile, string iUnZipPath, bool iCreateRelativePath, bool iOverWriteExistFile, string iPassword)
        public bool UnZipFile(string iZipFile, string iUnZipPath, bool iOverWriteExistFile, string iPassword)
        {
            bool Result = false;

            try
            {
                using (ZipFile Unzip = new ZipFile(iZipFile, System.Text.Encoding.Default))
                {
                    if (!string.IsNullOrEmpty(iPassword))
                        Unzip.Password = iPassword;

                    //將檔案各別解壓縮
                    for (int i = 0; i < Unzip.Count; i++) // (ZipEntry e in Unzip)
                    {
                        ZipEntry e = Unzip[i];

                       // if (!iCreateRelativePath)   //表示解壓縮時，不需要包含原始路徑
                       //     e.FileName = Path.GetFileName(e.FileName);

                        e.Extract(iUnZipPath, iOverWriteExistFile ? ExtractExistingFileAction.OverwriteSilently : ExtractExistingFileAction.DoNotOverwrite);
                    }
                }

                Result = true;
            }
            catch (Exception ex)
            {
                throw;
            }

            return Result;
        }
        #endregion
    }
}
