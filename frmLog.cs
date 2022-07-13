using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TimerExecute
{
    public partial class frmLog : Form
    {
        private List<string> sysLog;
        public frmLog(List<string> logs)
        {
            sysLog = logs;
            InitializeComponent();
        }

        private void frmLog_Load(object sender, EventArgs e)
        {
            foreach (string log in sysLog)
            {
                lsbLog.Items.Add(log);
            }
        }
    }
}
