using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SIC_Tool.Common;
using SIC_Tool.Common.Model;

namespace SIC_Tool.WinForm
{
    public partial class ChangePasswordForm : Form
    {
        public ChangePasswordForm()
        {
            InitializeComponent();
        }

        private void ChangePasswordButton_Click(object sender, EventArgs e)
        {
            string useridstring = "User Id=";
            string passwordstring = "Password=";
            string ntaccount = NTAccountTextBox.Text.Trim();
            string password = PasswordTextBox.Text.Trim();
            string confirmpassword = ConfirmPasswordTextBox.Text.Trim();

            if (NTAccountTextBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter your NT Account!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!password.Equals(confirmpassword))
            {
                MessageBox.Show("Your password doesn't match!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<AppSetting> connectionstringlist = new List<AppSetting>();
            connectionstringlist = ConfigFileUtility.GetKeyValueList();

            foreach (AppSetting connectionstring in connectionstringlist)
            {
                int startposition = 0;
                int endposition = 0;
                string existinguserid = string.Empty;
                string existingpassword = string.Empty;

                if (connectionstring.KeyValue.ToUpper().Contains(useridstring.ToUpper()))
                {
                    startposition = connectionstring.KeyValue.IndexOf(useridstring) + useridstring.Length;
                    endposition = connectionstring.KeyValue.IndexOf(passwordstring) - 1;
                    existinguserid = connectionstring.KeyValue.Substring(startposition, endposition - startposition);

                    if (connectionstring.KeyValue.ToUpper().Contains(passwordstring.ToUpper()))
                    {
                        startposition = connectionstring.KeyValue.IndexOf(passwordstring) + passwordstring.Length;
                        endposition = connectionstring.KeyValue.LastIndexOf(";");
                        existingpassword = connectionstring.KeyValue.Substring(startposition, endposition - startposition);

                        StringBuilder sb = new StringBuilder(connectionstring.KeyValue);
                        sb.Replace(existinguserid, ntaccount);
                        sb.Replace(existingpassword, PasswordUtility.DesEncrypt(password));

                        ConfigFileUtility.SetValue(connectionstring.Key, sb.ToString());
                    }
                }
            }

            SynchronizeConfiguration();

            MessageBox.Show("Your password has been changed successful!", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            NTAccountTextBox.Text = "";
            PasswordTextBox.Text = "";
            ConfirmPasswordTextBox.Text = "";
        }

        private void SynchronizeConfiguration()
        {
            string configfilename = ConfigFileUtility.GetValue("BackEndJobConfigFile");
            string configfilepath = Path.GetDirectoryName(configfilename);
            string newconfigfilename = Path.ChangeExtension(configfilename, ".bak");
            string fullconfigfilename = Path.Combine(Path.GetFullPath(configfilename), newconfigfilename);

            if (!File.Exists(configfilepath))
            {
                MessageBox.Show("The password will not be synchronized due to the configuration file of Glovia pulling job cannot be found", 
                    "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                if (File.Exists(newconfigfilename))
                    File.Delete(newconfigfilename);

                File.Move(configfilename, fullconfigfilename);

                string sourceconfigfilename = string.Format("{0}.config", System.Windows.Forms.Application.ExecutablePath);

                File.Copy(sourceconfigfilename, configfilename);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }    
}
