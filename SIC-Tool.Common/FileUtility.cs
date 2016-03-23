using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;
using SIC_Tool.Common.Model;

namespace SIC_Tool.Common
{
    public class FileUtility
    {
        private static ILog log = LogManager.GetLogger(typeof(FileUtility));

        public static bool IsValidTime(string fileName)
        {
            int waitTimeSpan = 10;
            FileInfo fi = new FileInfo(fileName);
            DateTime lastModified = fi.LastWriteTime;
            TimeSpan timeSpan = DateTime.Now - lastModified;

            if (timeSpan.TotalSeconds >= waitTimeSpan) return true;
            return false;
        }

        public static bool IsFileExisting(string filename, string targetfolder)
        {
            DirectoryInfo dir = new DirectoryInfo(targetfolder);

            foreach (FileInfo fi in dir.GetFiles())
            {
                if (filename == fi.Name) return true;
            }

            return false;
        }

        public static string LoadTextFile(string path)
        {
            string text = null;

            try
            {
                using (StreamReader reader = new FileInfo(path).OpenText())
                {
                    text = reader.ReadToEnd();
                }

                return text;
            }
            catch (FileNotFoundException ex)
            {
                throw new FileNotFoundException(
                    string.Format("The file {0} cannot be found: {1}", path, ex.Message));
            }
            catch (FileLoadException ex)
            {
                throw new FileLoadException(
                    string.Format("Loading the file {0} failed: {1}", path, ex.Message));
            }
            catch
            {
                throw;
            }
        }

        public static void MoveFile(string targefolder, FileInfo filename)
        {
            string newFileName = String.Empty;

            if (string.IsNullOrEmpty(targefolder))
            {
                throw new ArgumentNullException("The targe folder of file cannot be empty!");
            }

            if (filename == null)
            {
                throw new ArgumentNullException("File name cannot be empty!");
            }

            try
            {
                newFileName = string.Format("{0}_{1}{2}",
                    filename.Name.Substring(0, filename.Name.Length),
                    DateTime.Now.ToFileTime().ToString(),
                    filename.Extension);
                filename.MoveTo(Path.Combine(targefolder, newFileName));
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Moving the file - {0} failed: {1}", filename.Name, ex.Message));
            }
        }

        public static void MoveFile(string targetfolder, string fullfilename)
        {
            string newFileName = String.Empty;

            if (!Directory.Exists(targetfolder))
            {
                Directory.CreateDirectory(targetfolder);
            }

            if (fullfilename == null)
            {
                throw new ArgumentNullException("File name cannot be empty!");
            }

            try
            {
                string filename = Path.GetFileNameWithoutExtension(fullfilename);
                string fileextension = Path.GetExtension(fullfilename);
                string newfilename = string.Format("{0}_{1}{2}", filename, DateTime.Now.ToString("yyyyMMdd_HHmmss"), fileextension);

                File.Move(fullfilename, Path.Combine(targetfolder, newfilename));
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Moving the file - {0} failed: {1}", fullfilename, ex.Message));
            }
        }

        public static void MoveFile(string targetfolder, string fullfilename, string newfilename)
        {
            string newFileName = String.Empty;

            if (!Directory.Exists(targetfolder))
            {
                Directory.CreateDirectory(targetfolder);
            }

            if (fullfilename == null)
            {
                throw new ArgumentNullException("File name cannot be empty!");
            }

            try
            {
                //newFileName = string.Format("{0}_{1}", filename.Name, DateTime.Now.ToFileTime().ToString());
                //string filename = Path.GetFileName(fullfilename);
                File.Move(fullfilename, Path.Combine(targetfolder, newfilename));
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Moving the file - {0} failed: {1}", fullfilename, ex.Message));
            }
        }
        //public static ReportFile LoadReportFileConfig()
        //{
        //    string xml = LoadTextFile(@".\ReportFile.xml");
        //    return (ReportFile)SerializationUtility.DeSerialize(typeof(ReportFile), xml);
        //}

        public static ReportConfig LoadReportConfig()
        {
            string xml = LoadTextFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ReportConfig.xml"));

            return (ReportConfig)SerializationUtility.DeSerialize(typeof(ReportConfig), MiscUtility.Char2HTML(xml));
        }

        public static void SaveFile(string filename, string text)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filename, true))
                {
                    writer.WriteLine(text);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Saving the file - {0} failed: {1}", filename, ex.Message));
            }
        }

        public static List<string> GetFileNameList(string foldername)
        {
            string filename = string.Empty;
            string fullfilename = string.Empty;

            string samplefilename = "20141018.xls";
            List<string> filenamelist = new List<string>();

            DirectoryInfo dir = new DirectoryInfo(foldername);
            foreach (FileInfo fi in dir.GetFiles())
            {
                //If file name like 20140410.xls
                if ((fi.Name.Length <= samplefilename.Length) && fi.Name.Substring(0, 2).Equals("20"))
                {
                    fullfilename = ExcelFileUtility.SaveAsStandardFileFormat(fi.FullName);

                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("Dispatch P/N");
                    columnnamelist.Add("Usage P/N");
                    ExcelFileUtility.ChangeColumnDataType(fullfilename, "Usage Report", columnnamelist, "'");
                    
                    filename = Path.GetFileNameWithoutExtension(fullfilename);
                    string targetfilename = string.Format("SRIE_{0}", filename);
                    string newfullfilename = fullfilename.Replace(filename, targetfilename);
                    File.Move(fullfilename, newfullfilename);

                     filenamelist.Add(newfullfilename);

                    continue;
                }

                //Convert "SHIPMENT_SER_p1_p3_CHN_proc20130411.csv" file to standard excel file.
                if (fi.Name.ToUpper().Contains("SHIPMENT_SER"))
                {                    
                    fullfilename = ExcelFileUtility.SaveAsStandardFileFormat(fi.FullName);
                    filenamelist.Add(fullfilename);

                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("PTA Nr /LPTA Nr");
                    ExcelFileUtility.ChangeColumnDataType(fullfilename, "SHIPMENT_SER", columnnamelist, "'");
                    
                    continue;
                }

                //Change the data type of the column name - "Service order" in the returninglist excel file
                if (fi.Name.ToUpper().Contains("RETURNINGLIST"))
                {
                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("Shipping Date");
                    columnnamelist.Add("Service Order");
                    columnnamelist.Add("Part Qty");
                    ExcelFileUtility.ChangeColumnDataType(fi.FullName, "returninglist", columnnamelist, "'");
                }

                //Change the data type of the column name - "Date" in the receivinglist excel file
                if (fi.Name.ToUpper().Contains("RECEIVINGLIST"))
                {
                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("Date");
                    columnnamelist.Add("UID/LPTA Nr");
                    columnnamelist.Add("Quantity");
                    columnnamelist.Add("Service Order");

                    ExcelFileUtility.ChangeColumnDataType(fi.FullName, "receivinglist", columnnamelist, "'");
                }

                if (fi.Name.ToUpper().Contains("REPORT_TAML"))
                {
                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("Account code");

                    ExcelFileUtility.ChangeColumnDataType(fi.FullName, "IA_report", columnnamelist, "'");
                }

                //if (fi.Name.ToUpper().Contains("SRIE_"))
                //{
                //    List<string> columnnamelist = new List<string>();
                //    columnnamelist.Add("Dispatch P/N");
                //    columnnamelist.Add("Usage P/N");

                //    ExcelFileUtility.ChangeColumnDataType(fi.FullName, "Usage Report", columnnamelist, "'");                    
                //}

                if (fi.Name.ToUpper().Contains("RECEIVING REPORT"))
                {
                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("Order Number");
                    columnnamelist.Add("Po Number");

                    ExcelFileUtility.ChangeColumnDataType(fi.FullName, "Receiving Report", columnnamelist, "'");
                }

                if (fi.Name.ToUpper().Contains("UNI_PARTS"))
                {
                    List<string> columnnamelist = new List<string>();
                    columnnamelist.Add("UID");
                    columnnamelist.Add("Call Number");

                    ExcelFileUtility.ChangeColumnDataType(fi.FullName, "Receiving", columnnamelist, "'");
                }

                filenamelist.Add(fi.FullName);
            }

            return filenamelist;
        }
    }
}
