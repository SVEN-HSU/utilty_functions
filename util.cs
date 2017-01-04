using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Web;
using System.Net.Sockets;
using System.Threading;
using System.Security.Cryptography; //MD5
using System.Globalization; //日期補零
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using System.Xml;
using ABYSSDrvShared;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Management;

namespace UTILITY_FUNCTIONS
{
    public static class DirectorySize
    {
        public static long GetSize(this DirectoryInfo dirInfo)
        {
            Type tp = Type.GetTypeFromProgID("Scripting.FileSystemObject");
            object fso = Activator.CreateInstance(tp);
            object fd = tp.InvokeMember("GetFolder", BindingFlags.InvokeMethod, null, fso, new object[] { dirInfo.FullName });
            long ret = Convert.ToInt64(tp.InvokeMember("Size", BindingFlags.GetProperty, null, fd, null));
            Marshal.ReleaseComObject(fso);
            return ret;
        }
    }

    public static class utils
    {
        public static long hddRemainSpace(string dirPath)
        {
            DirectoryInfo di = new DirectoryInfo(dirPath);            
            DriveInfo drvInfo = new DriveInfo(di.Root.ToString());

            return drvInfo.AvailableFreeSpace;
        }

        public static string sizeFormat(object filesize)
        {
            string result = "";
            double i_result = 0;
            string err = "";
            try
            {
                i_result = Convert.ToDouble(filesize);
                if (((i_result / 1099511627776) > 1))
                {
                    result = (i_result / 1099511627776).ToString("F2") + " TB";
                    return result;                
                }
                if ((i_result / 1073741824) > 1)
                {
                    result = (i_result / 1073741824).ToString("F2") + " GB";
                    return result;
                }
                else if ((i_result / 1048576) > 1)
                {
                    result = (i_result / 1048576).ToString("F2") + " MB";

                    return result;

                }
                else if ((i_result / 1024) > 1)
                {
                    result = (i_result / 1024).ToString("F2") + " KB";
                    return result;
                }
                else
                {
                    result = i_result.ToString() + " Bytes";
                    return result;
                }
            }
            catch (Exception excep)
            {
                err = excep.Message;
            }
            return err;
        }

        public static long getDirectorySize(string folderPath)
        {
            DirectoryInfo di = new DirectoryInfo(folderPath);
            return di.GetSize();
        }

        public static int getProcessID(string processName, string keyValue, ref string errMessage)
        {
            ManagementClass mngmtClass = new ManagementClass("Win32_Process");
            int returnVal = 0;
            try
            {
                foreach (ManagementObject o in mngmtClass.GetInstances())
                {
                    if (o["Name"].Equals(processName) && Convert.ToString(o["CommandLine"]).Contains(keyValue))
                    {
                        returnVal = Convert.ToInt32(o["ProcessId"]);
                        break;
                    }
                } 
            }
            catch (Exception err)
            {
                errMessage = err.Message;
                return -1;
            }

            return returnVal;
        }

        public static string checkFileOpenedOrMissing(string filePath)
        {
            if (File.Exists(filePath) == false)
            {
                return "missingFile";
            }

            if (isFileOpened(filePath) == true)
            {
                return "fileOpened";
            }

            return "ok";
        }

        public static bool isFileOpened(string filepath)
        {
            LogHandle log = new LogHandle();
            logType LOGTYPE = logType.Upload;
            string LOGMODE = "1";
            string status = "ok";

            if (File.Exists(filepath))
            {
                FileStream stream = null;
                FileInfo fi = new FileInfo(filepath);
                try
                {
                    stream = fi.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                catch (IOException _ioerr)
                {
                    log.WriteToLog("[isFileOpened] IOException:" + _ioerr.Message, logType.Upload, logLevel.Exception, LOGMODE);
                    return true;
                }
                catch (UnauthorizedAccessException ua)
                {
                    log.WriteToLog("[isFileOpened] UnauthorizedAccessException:" + ua.Message, logType.Upload, logLevel.Exception, LOGMODE);
                    return true;
                }
                catch (Exception err)
                {
                    log.WriteToLog("[isFileOpened] Exception:" + err.Message, logType.Upload, logLevel.Exception, LOGMODE);
                    return true;
                }
                stream.Close();
                stream.Dispose();

                return false;
            }
            else
            {
                return true;
            }
        }

        public static string timeSpanFormat(double total_seconds)
        {
            TimeSpan t = TimeSpan.FromSeconds(total_seconds);
            string answer = "";
            if (total_seconds > 3600)
            {
                answer = string.Format("{0:D2}:{1:D2}:{2:D2}:", t.Hours, t.Minutes, t.Seconds);
            }
            else if (total_seconds > 60 && total_seconds < 3600)
            {
                answer = string.Format("{0:D2}:{1:D2}", t.Minutes, t.Seconds);
            }

            return answer;
        }
    }
}
