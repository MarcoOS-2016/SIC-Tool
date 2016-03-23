using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SIC_Tool.WinForm
{
    public partial class ProgressForm : Form
    {
        private bool cancelstatus = false;

        public bool CancelStatus
        {
            get { return cancelstatus; }
        }

        public ProgressForm()
        {
            InitializeComponent();
            pictureBox1.BackColor = Color.Transparent;
        }

        private void ProgressCancelButton_Click(object sender, EventArgs e)
        {
            cancelstatus = true;
            //e.Cancel = true;
        }
    }
}
