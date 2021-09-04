namespace WordReport
{
    using Microsoft.Win32;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Text;
    using System.Threading;
    using System.Windows.Forms;

    partial class MainForm : Form
    {
        private int mDocType;
        private string mTemplate = "";
        private bool mIncludeImage;
        private string mRecNo = "";
        private string mEaglePath = "";
        private string strReportName = string.Empty;
        private System.Windows.Forms.Timer timerShowWord;
        private int nTimerCount;
        private Dictionary<int, string> dictDocumentType;
        private IContainer components;
        private Button btnGenerate;
        private Label lblDocumentType;
        private Label lblTemplate;
        private Label IncludeImage;
        private Label lblRefNo;
        private ComboBox cmbDocumentType;
        private CheckBox cbIncludeImage;
        private TextBox textRefNo;
        private Button btnCancel;
        private ComboBox cmbTemplate;
        private Label label1;
        private TextBox textEaglePath;
        private GroupBox groupBox1;
        private Label lbProgress;
        private ProgressBar pbReport;

        public MainForm()
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            this.InitializeComponent();
            this.InitControls();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            this.UpdateData(true);
            new Thread(new ThreadStart(this.GenerateReport)).Start();
        }

        private bool CheckMDACVersion()
        {
            RegistryKey key = null;
            RegistryKey localMachine = Registry.LocalMachine;
            try
            {
                key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\DataAccess", RegistryKeyPermissionCheck.ReadSubTree);
            }
            catch (Exception exception1)
            {
                MessageBox.Show(exception1.Message);
                return true;
            }
            if (key == null)
            {
                return false;
            }
            Version version = new Version(key.GetValue("FullInstallVer", "2.50.0.0").ToString());
            return ((version.Major > 2) || (version.Minor >= 60));
        }

        private void cmbDocumentType_SelectedIndexChanged(object sender, EventArgs args)
        {
            this.InitTemplate();
            this.InitIncludeImage();
        }

        private EagleReport CreateReport(int nDocType, string strTemplate, bool bIncludeImage, string strRecNo, string strEaglePath)
        {
            EagleReport report = null;
            switch (nDocType)
            {
                case 0:
                    report = new ContractReport(strTemplate, bIncludeImage, strRecNo, strEaglePath);
                    break;

                case 3:
                    report = new PurchaseReport(strTemplate, bIncludeImage, strRecNo, strEaglePath);
                    break;

                case 4:
                    report = new NewQuotReport(strTemplate, bIncludeImage, strRecNo, strEaglePath);
                    break;

                default:
                    MessageBox.Show("The function hasnot been implemented !");
                    break;
            }
            return report;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void EnableInputControls(bool bEnabled)
        {
            this.cmbDocumentType.Enabled = bEnabled;
            this.cmbTemplate.Enabled = bEnabled;
            this.cbIncludeImage.Enabled = bEnabled;
            this.textRefNo.Enabled = bEnabled;
            this.textEaglePath.Enabled = bEnabled;
            this.btnGenerate.Enabled = bEnabled;
            this.btnCancel.Enabled = bEnabled;
        }

        private string FindArgument(string strPrefix)
        {
            foreach (string str in Environment.GetCommandLineArgs())
            {
                if (str.StartsWith(strPrefix, StringComparison.InvariantCultureIgnoreCase))
                {
                    return str;
                }
            }
            return null;
        }

        private bool FindWordWindow(IntPtr hWnd, int y)
        {
            StringBuilder text = new StringBuilder(0xff);
            User32Helper.GetWindowText(hWnd, text, 0xff);
            if (string.IsNullOrEmpty(this.strReportName) || !text.ToString().Contains(this.strReportName))
            {
                return true;
            }
            this.timerShowWord.Stop();
            this.report_ShowProgress(ReportStatus.RS_COMPLETEALL, 100);
            base.Close();
            return false;
        }

        private void GenerateReport()
        {
            EagleReport report = this.CreateReport(this.mDocType, this.mTemplate, this.mIncludeImage, this.mRecNo, this.mEaglePath);
            if (report != null)
            {
                report.ShowProgressEvent += new ShowProgressEvent(this.report_ShowProgress);
                report.Init();
                this.strReportName = report.Execute();
                if (!string.IsNullOrEmpty(this.strReportName))
                {
                    Process.Start(Path.Combine(EagleReport.GetReportPath(), this.strReportName));
                }
                else
                {
                    MessageBox.Show("\x00c9\x00fa\x00b3\x00c9\x00b1\x00a8\x00b8\x00e6\x00ca\x00b1\x00b7\x00a2\x00c9\x00fa\x00b4\x00ed\x00ce\x00f3\x00a3\x00a1");
                }
                report.ShowProgressEvent -= new ShowProgressEvent(this.report_ShowProgress);
            }
        }

        private void InitArgs()
        {
            string str = this.FindArgument("/DocType");
            if (str != null)
            {
                this.mDocType = int.Parse(str.Substring(str.IndexOf(":") + 1));
            }
            str = this.FindArgument("/Template");
            if (str != null)
            {
                this.mTemplate = str.Substring(str.IndexOf(":") + 1);
            }
            if (this.FindArgument("/IncludeImage") != null)
            {
                this.mIncludeImage = true;
            }
            str = this.FindArgument("/RecNo");
            if (str != null)
            {
                this.mRecNo = str.Substring(str.IndexOf(":") + 1);
            }
            str = this.FindArgument("/EaglePath");
            if (str != null)
            {
                this.mEaglePath = str.Substring(str.IndexOf(":") + 1);
            }
        }

        private void InitControls()
        {
            this.InitDocType();
            this.InitTemplate();
            this.InitIncludeImage();
        }

        private void InitDocType()
        {
            this.dictDocumentType = new Dictionary<int, string>();
            this.dictDocumentType.Add(0, "Sales Confirmation");
            this.dictDocumentType.Add(3, "Purchase Order");
            this.dictDocumentType.Add(4, "New Quotation");
            this.cmbDocumentType.Items.Clear();
            foreach (string str in this.dictDocumentType.Values)
            {
                this.cmbDocumentType.Items.Add(str);
            }
            this.cmbDocumentType.SelectedIndex = 0;
            this.cmbDocumentType.SelectedIndexChanged += new EventHandler(this.cmbDocumentType_SelectedIndexChanged);
        }

        private void InitializeComponent()
        {
            this.btnGenerate = new Button();
            this.lblDocumentType = new Label();
            this.lblTemplate = new Label();
            this.IncludeImage = new Label();
            this.lblRefNo = new Label();
            this.cmbDocumentType = new ComboBox();
            this.cbIncludeImage = new CheckBox();
            this.textRefNo = new TextBox();
            this.btnCancel = new Button();
            this.cmbTemplate = new ComboBox();
            this.label1 = new Label();
            this.textEaglePath = new TextBox();
            this.groupBox1 = new GroupBox();
            this.lbProgress = new Label();
            this.pbReport = new ProgressBar();
            this.groupBox1.SuspendLayout();
            base.SuspendLayout();
            this.btnGenerate.Location = new Point(0x2a, 0x124);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new Size(0x4b, 0x17);
            this.btnGenerate.TabIndex = 0;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new EventHandler(this.btnGenerate_Click);
            this.lblDocumentType.AutoSize = true;
            this.lblDocumentType.Location = new Point(0x13, 0x19);
            this.lblDocumentType.Name = "lblDocumentType";
            this.lblDocumentType.Size = new Size(0x53, 13);
            this.lblDocumentType.TabIndex = 2;
            this.lblDocumentType.Text = "DocumentType:";
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Location = new Point(0x13, 0x31);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Size = new Size(0x36, 13);
            this.lblTemplate.TabIndex = 3;
            this.lblTemplate.Text = "Template:";
            this.IncludeImage.AutoSize = true;
            this.IncludeImage.Location = new Point(0x13, 0x4b);
            this.IncludeImage.Name = "IncludeImage";
            this.IncludeImage.Size = new Size(0x4a, 13);
            this.IncludeImage.TabIndex = 4;
            this.IncludeImage.Text = "IncludeImage:";
            this.lblRefNo.AutoSize = true;
            this.lblRefNo.Location = new Point(0x13, 0x68);
            this.lblRefNo.Name = "lblRefNo";
            this.lblRefNo.Size = new Size(50, 13);
            this.lblRefNo.TabIndex = 5;
            this.lblRefNo.Text = "REF NO:";
            this.cmbDocumentType.FormattingEnabled = true;
            object[] items = new object[] { "SalesConfirmation", "Invoice", "Purchas Order" };
            this.cmbDocumentType.Items.AddRange(items);
            this.cmbDocumentType.Location = new Point(0x9e, 0x11);
            this.cmbDocumentType.Name = "cmbDocumentType";
            this.cmbDocumentType.Size = new Size(0x81, 0x15);
            this.cmbDocumentType.TabIndex = 6;
            this.cmbDocumentType.TabStop = false;
            this.cmbDocumentType.Text = "Sales Confirmation";
            this.cbIncludeImage.AutoSize = true;
            this.cbIncludeImage.Location = new Point(0x9e, 0x4b);
            this.cbIncludeImage.Name = "cbIncludeImage";
            this.cbIncludeImage.Size = new Size(15, 14);
            this.cbIncludeImage.TabIndex = 7;
            this.cbIncludeImage.UseVisualStyleBackColor = true;
            this.textRefNo.Location = new Point(0x9e, 0x61);
            this.textRefNo.Name = "textRefNo";
            this.textRefNo.Size = new Size(0x7f, 20);
            this.textRefNo.TabIndex = 8;
            this.btnCancel.Location = new Point(210, 0x124);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new Size(0x4b, 0x17);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new EventHandler(this.btnCancel_Click);
            this.cmbTemplate.FormattingEnabled = true;
            this.cmbTemplate.Location = new Point(0x9e, 0x2e);
            this.cmbTemplate.Name = "cmbTemplate";
            this.cmbTemplate.Size = new Size(0x7f, 0x15);
            this.cmbTemplate.TabIndex = 9;
            this.label1.AutoSize = true;
            this.label1.Location = new Point(0x13, 140);
            this.label1.Name = "label1";
            this.label1.Size = new Size(0x38, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "EaglePath";
            this.textEaglePath.Location = new Point(0x9e, 0x85);
            this.textEaglePath.Name = "textEaglePath";
            this.textEaglePath.Size = new Size(0x7f, 20);
            this.textEaglePath.TabIndex = 8;
            this.groupBox1.Controls.Add(this.lbProgress);
            this.groupBox1.Controls.Add(this.pbReport);
            this.groupBox1.Location = new Point(0x16, 0xb0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new Size(0x134, 0x61);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Progress";
            this.lbProgress.AutoSize = true;
            this.lbProgress.Location = new Point(0x11, 0x20);
            this.lbProgress.Name = "lbProgress";
            this.lbProgress.Size = new Size(0, 13);
            this.lbProgress.TabIndex = 13;
            this.pbReport.Location = new Point(6, 0x45);
            this.pbReport.Name = "pbReport";
            this.pbReport.Size = new Size(0x128, 0x16);
            this.pbReport.TabIndex = 12;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x156, 0x147);
            base.Controls.Add(this.groupBox1);
            base.Controls.Add(this.cmbTemplate);
            base.Controls.Add(this.textEaglePath);
            base.Controls.Add(this.textRefNo);
            base.Controls.Add(this.cbIncludeImage);
            base.Controls.Add(this.label1);
            base.Controls.Add(this.cmbDocumentType);
            base.Controls.Add(this.lblRefNo);
            base.Controls.Add(this.IncludeImage);
            base.Controls.Add(this.lblTemplate);
            base.Controls.Add(this.lblDocumentType);
            base.Controls.Add(this.btnCancel);
            base.Controls.Add(this.btnGenerate);
            this.Cursor = Cursors.WaitCursor;
            base.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            base.Name = "MainForm";
            base.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Word Report";
            base.UseWaitCursor = true;
            base.Load += new EventHandler(this.MainForm_Load);
            base.Shown += new EventHandler(this.MainForm_Shown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        private void InitIncludeImage()
        {
            this.cbIncludeImage.Checked = false;
        }

        private void InitTemplate()
        {
            int num = -1;
            foreach (int num2 in this.dictDocumentType.Keys)
            {
                if (this.dictDocumentType[num2] == this.cmbDocumentType.Text)
                {
                    num = num2;
                    break;
                }
            }
            Dictionary<int, string> dictionary = new Dictionary<int, string>();
            switch (num)
            {
                case 0:
                    dictionary.Add(0, "XMCONTRACT.DOC");
                    dictionary.Add(1, "HKCONTRACT.DOC");
                    break;

                case 1:
                    dictionary.Add(0, "XMINVOICE.DOC");
                    dictionary.Add(1, "HKINVOICE.DOC");
                    dictionary.Add(2, "CERINVOICE.DOC");
                    dictionary.Add(3, "BLANKINVOICE.DOC");
                    break;

                case 2:
                    dictionary.Add(0, "XMPACKINGLIST.DOC");
                    dictionary.Add(1, "HKPACKINGLIST.DOC");
                    dictionary.Add(2, "BJPACKINGLIST.DOC");
                    break;

                case 3:
                    dictionary.Add(0, "PURCHASE.DOC");
                    dictionary.Add(1, "BLANKPURCHASE.DOC");
                    break;

                case 4:
                    dictionary.Add(0, "NEWQUOTATION.DOC");
                    break;

                default:
                    MessageBox.Show("DocumentType Undefined!");
                    break;
            }
            this.cmbTemplate.Items.Clear();
            foreach (string str in dictionary.Values)
            {
                this.cmbTemplate.Items.Add(str);
            }
            this.cmbTemplate.SelectedIndex = 0;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            if (!this.CheckMDACVersion())
            {
                MessageBox.Show("You need to update Microsoft MDAC to 2.6 or above!");
                base.Close();
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            if (Environment.GetCommandLineArgs().Length > 1)
            {
                this.InitArgs();
                this.UpdateData(false);
                this.EnableInputControls(false);
                this.nTimerCount = 0;
                this.StartTimer();
                new Thread(new ThreadStart(this.GenerateReport)).Start();
            }
        }

        private void report_ShowProgress(ReportStatus reportStatus, int nProgressStep)
        {
            string reportStatusMessage = EagleReport.GetReportStatusMessage(reportStatus);
            this.lbProgress.Text = reportStatusMessage;
            this.lbProgress.Invalidate();
            this.pbReport.Value = nProgressStep;
        }

        private void StartTimer()
        {
            this.timerShowWord = new System.Windows.Forms.Timer();
            this.timerShowWord.Interval = 0x7d0;
            this.timerShowWord.Tick += new EventHandler(this.timerShowWord_Tick);
            this.timerShowWord.Start();
        }

        private void timerShowWord_Tick(object sender, EventArgs e)
        {
            this.nTimerCount++;
            if (this.nTimerCount >= 3)
            {
                this.timerShowWord.Stop();
                base.Close();
            }
            User32Helper.EnumWindows(new EnumWindowsEvent(this.FindWordWindow), 0);
        }

        private void UpdateData(bool bSaveAndValidate)
        {
            if (!bSaveAndValidate)
            {
                this.cmbDocumentType.Text = this.dictDocumentType[this.mDocType];
                this.cmbTemplate.Text = this.mTemplate;
                this.cbIncludeImage.Checked = this.mIncludeImage;
                this.textRefNo.Text = this.mRecNo;
                this.textEaglePath.Text = this.mEaglePath;
            }
            else
            {
                foreach (int num in this.dictDocumentType.Keys)
                {
                    if (this.dictDocumentType[num] == this.cmbDocumentType.Text)
                    {
                        this.mDocType = num;
                        break;
                    }
                }
                this.mTemplate = this.cmbTemplate.Text;
                this.mIncludeImage = this.cbIncludeImage.Checked;
                this.mRecNo = this.textRefNo.Text;
                this.mEaglePath = this.textEaglePath.Text;
            }
        }
    }
}
