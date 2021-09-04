namespace WordReport
{
    using System;
    using System.Data;
    using System.Data.Odbc;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Runtime.CompilerServices;

    internal abstract class EagleReport
    {
        private string mConnString;
        private string mTemplate;
        private bool mIncludeImage;
        private string mRecNo;
        private string mEaglePath;
        private ShowProgressEvent ShowProgress;

        internal event ShowProgressEvent ShowProgressEvent
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            add
            {
                this.ShowProgress += value;
            }
            [MethodImpl(MethodImplOptions.Synchronized)]
            remove
            {
                this.ShowProgress -= value;
            }
        }

        public EagleReport(string strTemplate, bool bIncludeImage, string strRecNo, string strEaglePath)
        {
            this.mTemplate = strTemplate;
            this.mIncludeImage = bIncludeImage;
            this.mRecNo = strRecNo;
            this.mEaglePath = strEaglePath;
        }

        public abstract string Execute();
        protected DataTable ExecuteDataTable(string commandText)
        {
            OdbcConnection connection = null;
            DataTable table2;
            try
            {
                connection = new OdbcConnection(this.ConnString);
                connection.Open();
                DataTable dataTable = new DataTable();
                new OdbcDataAdapter(new OdbcCommand(commandText, connection)).Fill(dataTable);
                table2 = dataTable;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return table2;
        }

        public Size GetAdjustedSize(Image image, int nWidth)
        {
            Size size = new Size();
            if (image.Width > (image.Height * 0.75))
            {
                size.Width = nWidth;
                size.Height = (image.Height * size.Width) / image.Width;
            }
            else
            {
                size.Height = (int)(nWidth * 0.75);
                size.Width = (size.Height * image.Width) / image.Height;
            }
            return size;
        }

        public string GetDocName(string strPrefix)
        {
            bool flag = true;
            string str = strPrefix + " " + this.mRecNo + ".doc";
            int num = 1;
            while (flag)
            {
                string path = Path.Combine(GetReportPath(), str);
                if (!File.Exists(path))
                {
                    flag = false;
                    continue;
                }
                string[] strArray = new string[] { strPrefix, " ", this.mRecNo, " ", num.ToString(), ".doc" };
                str = string.Concat(strArray);
                num++;
            }
            return str;
        }

        protected string GetEnglishDateTime(string strVal, string strFormat)
        {
            DateTime time = DateTime.Parse(strVal);
            CultureInfo provider = CultureInfo.CreateSpecificCulture("en-US");
            return time.ToString(strFormat, provider);
        }

        public static string GetReportPath()
        {
            if (!Directory.Exists(@"C:\EagleReport\DOC"))
            {
                try
                {
                    Directory.CreateDirectory(@"C:\EagleReport\DOC");
                }
                catch (Exception exception1)
                {
                    throw exception1;
                }
            }
            return @"C:\EagleReport\DOC";
        }

        public static string GetReportStatusMessage(ReportStatus reportStatus)
        {
            string str = null;
            switch (reportStatus)
            {
                case ReportStatus.RS_START:
                    str = "Start word report";
                    break;

                case ReportStatus.RS_COMPLETEMAINDOC:
                    str = "Main document info Complete";
                    break;

                case ReportStatus.RS_COMPLETEDETAIL:
                    str = "Detail info Complete";
                    break;

                case ReportStatus.RS_READTOMERGEDATA:
                    str = "Merging data to Word Template.";
                    break;

                case ReportStatus.RS_WAITFORWORDTOSTART:
                    str = "Merging completed. Wait for MsWord to start...";
                    break;

                case ReportStatus.RS_COMPLETEALL:
                    str = "Wordreport generated successfully";
                    break;

                default:
                    break;
            }
            return str;
        }

        internal void Init()
        {
            this.mConnString = "DSN=Eagle;UID=e;PWD=e";
        }

        protected void OnShowProgress(ReportStatus reportStatus, int nProgressStep)
        {
            if (this.ShowProgress != null)
            {
                this.ShowProgress(reportStatus, nProgressStep);
            }
        }

        protected string Template =>
            this.mTemplate;

        protected string RecNo =>
            this.mRecNo;

        protected bool IncludeImage =>
            this.mIncludeImage;

        protected string ConnString =>
            this.mConnString;

        protected string EaglePath =>
            this.mEaglePath;

        protected string EagleDocPath =>
            Path.Combine(this.mEaglePath, "DOT");

        protected string EagleImagePath =>
            Path.Combine(this.mEaglePath, "Image");

        protected string EagleNewImagePath =>
            Path.Combine(this.mEaglePath, "NewImage");
    }
}
