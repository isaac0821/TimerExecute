using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace TimerExecute
{
    public partial class frmMain : Form
    {
        private double frequencyExe;
        private bool cycleExe;
        private int remainExe;
        private DateTime? lastExe;
        private DateTime? nextExe;
        private string executeFilePath = "";
        private string executeFileName = "";
        private string executeFileExtension = "";
        private string executeFileDisk = "";
        private List<string> sysLog = new List<string>();
        private DateTime startAt = DateTime.Now;

        Timer curClockTimer = new Timer();

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            lblHour.Text = ((DateTime.Now.Hour <= 12 ? DateTime.Now.Hour : (DateTime.Now.Hour - 12)) < 10 ? "0" : "")
                + (DateTime.Now.Hour <= 12 ? DateTime.Now.Hour.ToString() : (DateTime.Now.Hour - 12).ToString());
            lblMinute.Text = (DateTime.Now.Minute < 10 ? "0" : "") + DateTime.Now.Minute.ToString();
            lblHalfDay.Text = DateTime.Now.Hour < 12 ? "AM" : "PM";
            lblHalfDay.ForeColor = DateTime.Now.Hour < 12 ? Color.Black : Color.Red;

            sysLog.Add(DateTime.Now.ToString() + ": " + "Welcome!");
            lblStatus.Text = "Idle";
            frequencyExe = 1;
            cycleExe = true;
            lastExe = null;
            nextExe = null;

            nicMain.Visible = false;

            // 每一分钟绘制一次当前的时刻
            TimeSpan secToNextMin = new TimeSpan();
            DateTime datetimeWithoutSec = new DateTime(
                DateTime.Now.Year,
                DateTime.Now.Month,
                DateTime.Now.Day,
                DateTime.Now.Hour,
                DateTime.Now.Minute,
                0);
            secToNextMin = DateTime.Now - datetimeWithoutSec;
            curClockTimer.Interval = (int)(1000 * (60 - secToNextMin.TotalSeconds));
            curClockTimer.Start();
            curClockTimer.Tick += RefreshClock;
        }
        int fixClock = 0;
        private void RefreshClock(object sender, EventArgs e)
        {
            curClockTimer.Interval = 1000 * 60;
            lblHour.Text = ((DateTime.Now.Hour <= 12 ? DateTime.Now.Hour : (DateTime.Now.Hour - 12)) < 10 ? "0" : "")
                + (DateTime.Now.Hour <= 12 ? DateTime.Now.Hour.ToString() : (DateTime.Now.Hour - 12).ToString());
            lblMinute.Text = (DateTime.Now.Minute < 10 ? "0" : "") + DateTime.Now.Minute.ToString();
            lblHalfDay.Text = DateTime.Now.Hour < 12 ? "AM" : "PM";
            lblHalfDay.ForeColor = DateTime.Now.Hour < 12 ? Color.Black : Color.Red;

            if ((cycleExe || remainExe > 0) && frequencyExe > 0 && nextExe != null && lblStatus.Text == "Waiting")
            {
                if (DateTime.Now.Year == (nextExe ?? new DateTime(9999, 12, 31)).Year &&
                    DateTime.Now.Month == (nextExe ?? new DateTime(9999, 12, 31)).Month &&
                    DateTime.Now.Day == (nextExe ?? new DateTime(9999, 12, 31)).Day &&
                    DateTime.Now.Hour == (nextExe ?? new DateTime(9999, 12, 31)).Hour &&
                    DateTime.Now.Minute == (nextExe ?? new DateTime(9999, 12, 31)).Minute)
                {
                    lastExe = DateTime.Now;
                    lblLastExe.Text = DateTime.Now.ToString();
                    lblLastExe.ForeColor = Color.Orange;
                    lblStatus.Text = "Executing";

                    ExecuteFile();

                    if (!cycleExe)
                    {
                        remainExe -= 1;
                        txtRemain.Text = remainExe.ToString();
                    }

                    if (cycleExe || remainExe > 0)
                    {
                        lblStatus.Text = "Waiting";
                    }
                    else
                    {
                        lblStatus.Text = "Idle";
                    }

                    FindNextExecute();
                }
            }
        }

        private void rdoHalfHour_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoHalfHour.Checked)
            {
                frequencyExe = 0.5;
                FindNextExecute();
            }
        }

        private void rdoOneHour_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoOneHour.Checked)
            {
                frequencyExe = 1;
                FindNextExecute();
            }
        }

        private void rdoTwoHour_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoTwoHour.Checked)
            {
                frequencyExe = 2;
                FindNextExecute();
            }
        }

        private void rdoOther_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoOther.Checked)
            {
                try
                {
                    frequencyExe = Convert.ToDouble(txtFreq.Text);
                    FindNextExecute();
                    txtFreq.Enabled = true;
                }
                catch (Exception)
                {
                    lblNextExe.Text = "(No Schedule)";
                    lblNextExe.ForeColor = Color.Black;
                    nextExe = null;
                }
            }
            else
            {
                txtFreq.Enabled = false;
            }
        }

        private void txtFreq_TextChanged(object sender, EventArgs e)
        {
            try
            {
                frequencyExe = Convert.ToDouble(txtFreq.Text);
                FindNextExecute();
            }
            catch (Exception)
            {
                lblNextExe.Text = "(No Schedule)";
                lblNextExe.ForeColor = Color.Black;
                nextExe = null;
            }
        }

        private void rdoCycle_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoCycle.Checked)
            {
                cycleExe = true;
                FindNextExecute();
            }
            else
            {
                cycleExe = false;
            }
        }

        private void rdoRemain_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoRemain.Checked)
            {
                cycleExe = false;
                remainExe = Convert.ToInt32(txtRemain.Text);
                FindNextExecute();
                txtRemain.Enabled = true;
            }
            else
            {
                cycleExe = true;
                txtRemain.Enabled = false;
            }
        }

        private void txtRemain_TextChanged(object sender, EventArgs e)
        {
            remainExe = Convert.ToInt32(txtRemain.Text);
            FindNextExecute();
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            OpenFileDialog findExe = new OpenFileDialog();
            findExe.Title = "Path to execute file";
            findExe.InitialDirectory = @"D:\";
            findExe.RestoreDirectory = true;
            if (findExe.ShowDialog() == DialogResult.OK)
            {
                lblPathExe.Text = Path.GetFullPath(findExe.FileName);
                executeFilePath = Path.GetDirectoryName(findExe.FileName);
                executeFileName = Path.GetFileName(findExe.FileName);
                executeFileExtension = Path.GetExtension(findExe.FileName);
                executeFileDisk = Path.GetPathRoot(findExe.FileName);
                lblStatus.Text = "Loaded";
                sysLog.Add(DateTime.Now.ToString() + ": " + "Load script " + lblPathExe.Text);
                FindNextExecute();
            }
        }

        private void ExecuteFile()
        {
            // 执行py文件
            if (executeFileExtension == ".py")
            {
                sysLog.Add(DateTime.Now.ToString() + ": " + "Start executing " + lblPathExe.Text);
                try
                {
                    Process cmd = new Process();
                    cmd.StartInfo.FileName = "cmd.exe";
                    cmd.StartInfo.UseShellExecute = false;
                    cmd.StartInfo.RedirectStandardInput = true;
                    cmd.StartInfo.RedirectStandardOutput = true;
                    cmd.StartInfo.RedirectStandardError = true;
                    cmd.StartInfo.CreateNoWindow = true;
                    cmd.Start();

                    cmd.StandardInput.WriteLine(executeFileDisk);
                    cmd.StandardInput.WriteLine("cd " + executeFilePath);
                    cmd.StandardInput.WriteLine("python " + executeFileName);
                    cmd.StandardInput.WriteLine("exit");

                    StreamReader output = cmd.StandardOutput;
                    string line = "";
                    int num = 1;
                    while ((line = output.ReadLine()) != null)
                    {
                        if (line != "")
                        {
                            sysLog.Add("Line " + num.ToString() + ": " + line);
                            num++;
                        }
                    }
                    cmd.WaitForExit();
                    cmd.Close();

                    lastExe = DateTime.Now;
                    lblLastExe.Text = lastExe.ToString();
                    lblLastExe.ForeColor = Color.Green;

                    sysLog.Add(DateTime.Now.ToString() + ": " + "Finished executing " + lblPathExe.Text);

                    int tipShowMilliseconds = 5000;
                    string tipTitle = "Execute succeed";
                    string tipContent = "Finished executing " + lblPathExe.Text;
                    ToolTipIcon tipType = ToolTipIcon.Info;
                    nicMain.ShowBalloonTip(tipShowMilliseconds, tipTitle, tipContent, tipType);
                }
                catch (Exception)
                {
                    lblLastExe.ForeColor = Color.Red;
                    sysLog.Add(DateTime.Now.ToString() + ": " + "Failed executing " + lblPathExe.Text);
                    int tipShowMilliseconds = 5000;
                    string tipTitle = "Execute Failed";
                    string tipContent = "Failed executing " + lblPathExe.Text;
                    ToolTipIcon tipType = ToolTipIcon.Error;
                    nicMain.ShowBalloonTip(tipShowMilliseconds, tipTitle, tipContent, tipType);
                }
            }
            else
            {
                MessageBox.Show("Stay tune for non-python script support");
            }
        }

        private void btnNow_Click(object sender, EventArgs e)
        {
            ExecuteFile();
        }

        private void lblStatus_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Author: Lan Peng");
        }

        private void lblStatus_TextChanged(object sender, EventArgs e)
        {
            if (lblStatus.Text == "Idle")
            {
                lblStatus.ForeColor = Color.Gray;
                lastExe = null;
                nextExe = null;
                btnNow.Enabled = false;
                dtpStart.Enabled = false;
                dtpStart.Value = DateTime.Now;
                btnStart.Enabled = false;
                btnStop.Enabled = false;
                btnFind.Enabled = true;
                rdoCycle.Enabled = false;
                rdoRemain.Enabled = false;
                rdoHalfHour.Enabled = false;
                rdoOneHour.Enabled = false;
                rdoTwoHour.Enabled = false;
                rdoOther.Enabled = false;
                lblLastExe.Text = "(No Record)";
                lblNextExe.Text = "(No Schedule)";
                lblPathExe.Text = "(Execute Path)";
                lblLastExe.ForeColor = Color.Black;
                lblNextExe.ForeColor = Color.Black;
            }
            else if (lblStatus.Text == "Loaded")
            {
                lblStatus.ForeColor = Color.Black;
                btnNow.Enabled = true;
                btnStart.Enabled = true;
                dtpStart.Enabled = true;
                rdoCycle.Enabled = true;
                rdoRemain.Enabled = true;
                rdoHalfHour.Enabled = true;
                rdoOneHour.Enabled = true;
                rdoTwoHour.Enabled = true;
                rdoOther.Enabled = true;
                if (nextExe != null)
                {
                    lblNextExe.ForeColor = Color.Red;
                }
                else
                {
                    lblNextExe.ForeColor = Color.Black;
                }
            }
            else if (lblStatus.Text == "Executing")
            {
                lblStatus.ForeColor = Color.Orange;
                btnFind.Enabled = false;
                dtpStart.Enabled = false;
                rdoCycle.Enabled = false;
                rdoRemain.Enabled = false;
                rdoHalfHour.Enabled = false;
                rdoOneHour.Enabled = false;
                rdoTwoHour.Enabled = false;
                rdoOther.Enabled = false;
                if (nextExe != null)
                {
                    lblNextExe.ForeColor = Color.Green;
                }
                else
                {
                    lblNextExe.ForeColor = Color.Black;
                }
            }
            else if (lblStatus.Text == "Waiting")
            {
                lblStatus.ForeColor = Color.Blue;
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                btnFind.Enabled = false;
                dtpStart.Enabled = false;
                rdoCycle.Enabled = false;
                rdoRemain.Enabled = false;
                rdoHalfHour.Enabled = false;
                rdoOneHour.Enabled = false;
                rdoTwoHour.Enabled = false;
                rdoOther.Enabled = false;
                if (nextExe != null)
                {
                    lblNextExe.ForeColor = Color.Green;
                }
                else
                {
                    lblNextExe.ForeColor = Color.Black;
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Waiting";

        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Idle";
        }

        private void btnLog_Click(object sender, EventArgs e)
        {
            frmLog frmLog = new frmLog(sysLog);
            frmLog.Show();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Loaded";
        }

        private void dtpStart_ValueChanged(object sender, EventArgs e)
        {
            startAt = dtpStart.Value;
            FindNextExecute();
        }

        private void FindNextExecute()
        {
            int intTHour = (int)(frequencyExe % 1);
            int intTMinute = (int)((frequencyExe - intTHour) * 60);
            if ((cycleExe || remainExe > 0) && frequencyExe > 0)
            {
                nextExe = startAt;
                while (nextExe < DateTime.Now)
                {
                    nextExe += new TimeSpan(intTHour, intTMinute, 0);
                }
                lblNextExe.Text = nextExe.ToString();
                if (lblStatus.Text == "Loaded")
                {
                    lblNextExe.ForeColor = Color.Red;
                }
                else if (lblStatus.Text == "Waiting")
                {
                    lblNextExe.ForeColor = Color.Green;
                }
            }
            else
            {
                nextExe = null;
                lblNextExe.Text = "(No Schedule)";
                lblNextExe.ForeColor = Color.Black;
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            sysLog.Add(DateTime.Now.ToString() + ": " + "Goodbye!");
            FileStream f = new FileStream("log.txt", FileMode.Append);
            foreach (string log in sysLog)
            {
                byte[] data = System.Text.Encoding.Default.GetBytes(log + "\r\n");
                f.Write(data, 0, data.Length);
            }
            f.Flush();
            f.Close();
            nicMain.Dispose();
            Dispose();
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.ShowInTaskbar = false;
                nicMain.Visible = true;
            }
        }

        private void nicMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                nicMain.Visible = false;
                this.ShowInTaskbar = true;
            }
        }
    }
}
