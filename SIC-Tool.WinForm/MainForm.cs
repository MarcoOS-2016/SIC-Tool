using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;
using SIC_Tool.Business;
using SIC_Tool.Business.ReportConversion;

namespace SIC_Tool.WinForm
{
    public partial class MainForm : Form
    {
        public static readonly ReportConfig reportconfig = FileUtility.LoadReportConfig();
        private SICContext SIC_Context = new SICContext();
        private bool IsError = false;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            ReportConversionSourceReportTextBox.Text = ConfigFileUtility.GetValue("Report_Conversion_Source_Folder");
            ReportConversionOutputFolderTextBox.Text = ConfigFileUtility.GetValue("Report_Conversion_Output_Folder");
            PullGloviaReportSavePathTextBox.Text = ConfigFileUtility.GetValue("Pull_Glovia_Report_Output_Folder");
            FBReportOutPutFolderTextBox.Text = ConfigFileUtility.GetValue("FB_Report_Output_Folder");
            FBReportSourceTextBox.Text = ConfigFileUtility.GetValue("FB_Report_FB_Original_Report");
            FBReportFBSCTextBox.Text = ConfigFileUtility.GetValue("FBSC_Report_File");
            EndingBalanceReportFolderTextBox.Text = ConfigFileUtility.GetValue("Ending_Balance_Report_Folder");
            EndingBalanceFBSCTextBox.Text = ConfigFileUtility.GetValue("FBSC_Report_File");
            EndingBalanceOutputFolderTextBox.Text = ConfigFileUtility.GetValue("Ending_Balance_Output_Folder");
        }
        #region ------------------------------ Backgroundworker ------------------------------
        ProgressForm progressform = null;

        public void StartSynchronizedJob(object instance)
        {
            // Create a background thread
            BackgroundWorker backgroundworker = new BackgroundWorker();
            backgroundworker.WorkerSupportsCancellation = true;
            backgroundworker.WorkerReportsProgress = true;
            backgroundworker.DoWork += new DoWorkEventHandler(DoWork);
            backgroundworker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(RunWorkerCompleted);
            backgroundworker.ProgressChanged += new ProgressChangedEventHandler(ProgressChanged);

            // Create a progress form on the UI thread
            progressform = new ProgressForm();

            // Kick off the Async thread
            backgroundworker.RunWorkerAsync(instance);

            // Lock up the UI with this modal progress form.
            progressform.ShowDialog(this);
            progressform = null;
        }

        private void DoWork(object sender, DoWorkEventArgs e)
        {
            string functionname = (string)e.Argument;

            switch (functionname)
            {
                case "ReportFileConversion":
                    ReportFileConversion();
                    break;

                case "PullGloviaReport":
                    PullGloviaReport();
                    break;

                case "ProcessFBOriginalReport":
                    ProcessFBOriginalReport();
                    break;

                case "ProcessEndingBalanceReport":
                    ProcessEndingBalanceReport();
                    break;

                case "ProcessLongAgingReport":
                    ProcessLongAgingReport();
                    break;
            }

            if (progressform.CancelStatus)
            {
                e.Cancel = true;
                return;
            }
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        private void RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // The background process is complete. First we should hide the
            // modal Progress Form to unlock the UI. The we need to inspect our
            // response to see if an error occured, a cancel was requested or
            // if we completed succesfully.

            // Hide the Progress Form
            if (progressform != null)
            {
                progressform.Hide();
                progressform = null;
            }

            // Check to see if an error occured in the 
            // background process.
            if (e.Error != null)
            {
                //IsError = true;
                MessageBox.Show(e.Error.Message);
                return;
            }

            // Check to see if the background process was cancelled.
            if (e.Cancelled)
            {
                MessageBox.Show("Processing cancelled!", "Cancel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Everything completed normally.
            // process the response using e.Result            
            //MessageBox.Show("Processing is complete!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        private void ReportConversionStartButton_Click(object sender, EventArgs e)
        {
            string sourcereportfolder = ReportConversionSourceReportTextBox.Text.Trim();
            string outputfolder = ReportConversionOutputFolderTextBox.Text.Trim();
            
            if (sourcereportfolder.Length == 0)
            {
                MessageBox.Show("Please select a source folder of report files", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputfolder.Length == 0)
            {
                MessageBox.Show("Please select a output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(sourcereportfolder))
            {
                MessageBox.Show("The Source Folder cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(outputfolder))
            {
                MessageBox.Show("The Output Folder cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SIC_Context.VendorReportFileSourceFolder = sourcereportfolder;
            SIC_Context.VendorReportFileOutputFolder = outputfolder;
            
            //if (SIC_Context.ReportFileNameList.Count < Enumeration.Report_File_Count)  //Total 8 3PL files
            //{
            //    MessageBox.Show("Please make sure that 8 3PL files are placed into the folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            try
            {
                StartSynchronizedJob("ReportFileConversion");

                if (!IsError)
                {
                    MessageBox.Show("Report Conversion completed!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripStatusLabel1.Text = "Done!";
                }

                dataGridView1.DataSource = SIC_Tool.Business.InFileHandler.ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void PullReportReportStartButton_Click(object sender, EventArgs e)
        {
            if (PullGloviaReportNameComboBox.SelectedItem.ToString().Length == 0)
            {
                MessageBox.Show("Please select a Report Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (PullGloviaReportCCNTextBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter CCN", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (PullGloviaReportMaslocTextBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter Masloc", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (PullGloviaReportSavePathTextBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please select a output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(PullGloviaReportSavePathTextBox.Text))
            {
                MessageBox.Show("The Output Folder cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SIC_Context.UIReport = new UIReport();
            SIC_Context.UIReport.ReportName = PullGloviaReportNameComboBox.SelectedItem.ToString().Trim().ToUpper();
            SIC_Context.UIReport.SnapShotDate = PullGloviaReportDateTimePicker.Value;
            SIC_Context.UIReport.CCN = PullGloviaReportCCNTextBox.Text.Trim().ToUpper();
            SIC_Context.UIReport.Masloc = PullGloviaReportMaslocTextBox.Text.Trim().ToUpper();
            SIC_Context.UIReport.SaveFolder = PullGloviaReportSavePathTextBox.Text.Trim().ToUpper();

            try
            {
                StartSynchronizedJob("PullGloviaReport");

                if (!IsError)
                {
                    MessageBox.Show("Glovia Report has been pull out!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripStatusLabel1.Text = "Done!";
                }                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void PullGloviaReportNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ccn = string.Empty;
            string masloc = string.Empty;

            if (PullGloviaReportNameComboBox.SelectedItem.ToString().Contains("Stock"))
            {
                ccn = ConfigFileUtility.GetValue("FBSC_CCN");
                masloc = ConfigFileUtility.GetValue("FBSC_Masloc");

                PullGloviaReportBeginDateTimePicker.Enabled = false;
                PullGloviaReportEndDateTimePicker.Enabled = false;
            }
            if (PullGloviaReportNameComboBox.SelectedItem.ToString().Contains("Transaction History"))
            {
                ccn = ConfigFileUtility.GetValue("TransactionHistoryReport_CCN");
                masloc = ConfigFileUtility.GetValue("TransactionHistoryREport_Masloc");

                PullGloviaReportBeginDateTimePicker.Enabled = true;
                PullGloviaReportEndDateTimePicker.Enabled = true;
            }

            PullGloviaReportCCNTextBox.Text = ccn;
            PullGloviaReportMaslocTextBox.Text = masloc;
        }

        private void FBReportStartButton_Click(object sender, EventArgs e)
        {
            string fboriginalreportpath = FBReportSourceTextBox.Text.Trim();
            string outputfolder = FBReportOutPutFolderTextBox.Text.Trim();
            string FBSCfilepath = FBReportFBSCTextBox.Text.Trim();
            DateTime cutoffdate = CutOffDateTimePicker.Value;

            if (cutoffdate.CompareTo(DateTime.Now) > 0)
            {
                MessageBox.Show("The date you select is later than current date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (fboriginalreportpath.Length == 0)
            {
                MessageBox.Show("Please select a report file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputfolder.Length == 0)
            {
                MessageBox.Show("Please select a output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (FBSCfilepath.Length == 0)
            {
                MessageBox.Show("Please select a stock status file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(fboriginalreportpath))
            {
                MessageBox.Show("The FB Original report cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(FBSCfilepath))
            {
                MessageBox.Show("The FBSC report cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SIC_Context.ReportFilePath = fboriginalreportpath;
            SIC_Context.OutPutFolder = outputfolder;
            SIC_Context.FBSCReporFilePath = FBSCfilepath;
            SIC_Context.CutOffDate = cutoffdate;

            try
            {
                StartSynchronizedJob("ProcessFBOriginalReport");

                if (!IsError)
                {
                    MessageBox.Show("FBReport has been processed successful!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripStatusLabel1.Text = "Done!";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void EndingBalanceStartButton_Click(object sender, EventArgs e)
        {
            string endingbalancereportfolder = EndingBalanceReportFolderTextBox.Text.Trim();
            string outputfolder = EndingBalanceOutputFolderTextBox.Text.Trim();
            string FBSCfilepath = EndingBalanceFBSCTextBox.Text.Trim();

            if (endingbalancereportfolder.Length == 0)
            {
                MessageBox.Show("Please select a report file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputfolder.Length == 0)
            {
                MessageBox.Show("Please select a output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (FBSCfilepath.Length == 0)
            {
                MessageBox.Show("Please select a stock status file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(endingbalancereportfolder))
            {
                MessageBox.Show("The Ending Balance Report Folder cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(outputfolder))
            {
                MessageBox.Show("The Output Folder cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(FBSCfilepath))
            {
                MessageBox.Show("The FBSC report cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SIC_Context.SourceFolder = endingbalancereportfolder;
            SIC_Context.OutPutFolder = outputfolder;
            SIC_Context.FBSCReporFilePath = FBSCfilepath;

            try
            {
                StartSynchronizedJob("ProcessEndingBalanceReport");

                if (!IsError)
                {
                    MessageBox.Show("Ending Balance Report has been processed successful!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripStatusLabel1.Text = "Done!";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ReportConversionSourceReportButton_Click(object sender, EventArgs e)
        {
            ReportConversionSourceReportTextBox.Text = SelectFolder();
        }

        private void ReportConversionOutputFolderBtton_Click(object sender, EventArgs e)
        {
            ReportConversionOutputFolderTextBox.Text = SelectFolder();
        }

        private void PullGloviaReportSavePathButton_Click(object sender, EventArgs e)
        {
            PullGloviaReportSavePathTextBox.Text = SelectFolder();
        }

        private void FBReportSourceButton_Click(object sender, EventArgs e)
        {
            FBReportSourceTextBox.Text = SelectFile();
        }
        
        private void FBReportFBSCButton_Click(object sender, EventArgs e)
        {
            FBReportFBSCTextBox.Text = SelectFile();
        }

        private void FBReportOutPutFolderButton_Click(object sender, EventArgs e)
        {
            FBReportOutPutFolderTextBox.Text = SelectFolder();
        }

        private void ChangePasswordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangePasswordForm passwordchangeform = new ChangePasswordForm();
            passwordchangeform.ShowDialog();
        }

        private void EndingBalanceSourceButton_Click(object sender, EventArgs e)
        {
            EndingBalanceReportFolderTextBox.Text = SelectFolder();
        }

        private void EndingBalanceStockStatusButton_Click(object sender, EventArgs e)
        {
            EndingBalanceFBSCTextBox.Text = SelectFile();
        }

        private void EndingBalanceOutputButton_Click(object sender, EventArgs e)
        {
            EndingBalanceOutputFolderTextBox.Text = SelectFolder();
        }

        private void FBReportStockStatusReportButton_Click(object sender, EventArgs e)
        {
            FBReportFBSCTextBox.Text = SelectFile();
        }
        #region ------------------------------ Common Code ------------------------------
        private string SelectFolder()
        {
            FolderBrowserDialog folderbrowser = new FolderBrowserDialog();
            folderbrowser.RootFolder = Environment.SpecialFolder.MyComputer;
            folderbrowser.SelectedPath = @"C:\";
            folderbrowser.ShowNewFolderButton = true;

            if (folderbrowser.ShowDialog() == DialogResult.OK)
            {
                return folderbrowser.SelectedPath;
            }

            return String.Empty;
        }

        private string SelectFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = @"C:\";
            openFileDialog.Filter = "Excel file (*.xls, *.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }

            return String.Empty;
        }
        #endregion

        private void ReportFileConversion()
        {
            try
            {
                InFileHandler infilehandler = new InFileHandler(SIC_Context);
                infilehandler.Process();
            }
            catch (Exception ex)
            {
                IsError = true;
                MessageBox.Show(
                    string.Format("Function name:[Report File Conversion] - Message:{0}, Source:{1}, StackTrack:{2}", 
                    ex.InnerException.Message, ex.Source, ex.StackTrace));
            }
        }

        private void PullGloviaReport()
        {
            try
            {
                UIInqueryReportHandler handler = new UIInqueryReportHandler(SIC_Context);
                handler.Process();
            }
            catch (Exception ex)
            {
                IsError = true;
                MessageBox.Show(
                    string.Format("Function name:[Pull Glovia Report] - Message:{0}, Source:{1}, StackTrack:{2}",
                    ex.InnerException.Message, ex.Source, ex.StackTrace));
            }
        }

        private void ProcessFBOriginalReport()
        {
            try
            {
                ReportConversionBase handler = new FBVarianceReportHandler(SIC_Context);
                handler.Process();
            }
            catch (Exception ex)
            {
                IsError = true;
                MessageBox.Show(
                    string.Format("Function name:[Process FB Report] - Message:{0}, Source:{1}, StackTrack:{2}",
                    ex.InnerException.Message, ex.Source, ex.StackTrace));
            }
        }

        private void ProcessEndingBalanceReport()
        {
            try
            {
                ReportConversionBase handler = new EndingBalanceReportHandler(SIC_Context);
                handler.Process();
            }
            catch (Exception ex)
            {
                IsError = true;
                MessageBox.Show(
                    string.Format("Function name:[Process EndingBalance Report] - Message:{0}, Source:{1}, StackTrack:{2}",
                    ex.InnerException.Message, ex.Source, ex.StackTrace));
            }
        }

        private void ProcessLongAgingReport()
        {
            //try
            //{
                ReportConversionBase handler = new LongAgingReportHandler(SIC_Context);
                handler.Process();
            //}
            //catch (Exception ex)
            //{
            //    IsError = true;
            //    MessageBox.Show(
            //        string.Format("Function name:[Process LongAging Report] - Message:{0}, Source:{1}, StackTrack:{2}",
            //        ex.InnerException.Message, ex.Source, ex.StackTrace));
            //}
        }

        private void LongAgingReportOpenButton_Click(object sender, EventArgs e)
        {
            LongAgingReportTextBox.Text = SelectFile();
        }

        private void FBSCReportOpenButton_Click(object sender, EventArgs e)
        {
            FBSCReportTextBox.Text = SelectFile();
        }

        private void OutputFolderOpenButton_Click(object sender, EventArgs e)
        {
            LongAgingReportOutputFolderTextBox.Text = SelectFolder();
        }

        private void LongAgingReportStartButton_Click(object sender, EventArgs e)
        {
            string longAgingReport = LongAgingReportTextBox.Text.Trim();
            string outputFolder = LongAgingReportOutputFolderTextBox.Text.Trim();
            string FBSCfilepath = FBSCReportTextBox.Text.Trim();

            if (longAgingReport.Length == 0)
            {
                MessageBox.Show("Please select a Long Aging report file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputFolder.Length == 0)
            {
                MessageBox.Show("Please select a output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (FBSCfilepath.Length == 0)
            {
                MessageBox.Show("Please select a stock status file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                MessageBox.Show("The Output Folder cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(FBSCfilepath))
            {
                MessageBox.Show("The FBSC report cannot be found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SIC_Context.ReportFilePath = longAgingReport;
            SIC_Context.OutPutFolder = outputFolder;
            SIC_Context.FBSCReporFilePath = FBSCfilepath;

            try
            {
                toolStripStatusLabel1.Text = string.Format("[0] - Start to generate Long Aging report...", DateTime.Now.ToString());
                StartSynchronizedJob("ProcessLongAgingReport");

                if (!IsError)
                {
                    MessageBox.Show("Long Aging Report has been processed successful!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    toolStripStatusLabel1.Text = string.Format("[0] - Done!", DateTime.Now.ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
