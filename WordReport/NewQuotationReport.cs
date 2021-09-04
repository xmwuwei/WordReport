namespace WordReport
{
    using Aspose.Words;
    using Aspose.Words.Reporting;
    using System;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;

    internal class NewQuotReport : EagleReport
    {
        private DocumentBuilder mBuilder;
        private double mRatio;
        private string mQuot_no;
        private bool mCbm;

        public NewQuotReport(string strTemplate, bool bIncludeImage, string strRecNo, string strEaglePath) : base(strTemplate, bIncludeImage, strRecNo, strEaglePath)
        {
            this.mRatio = 1.0;
        }

        public override string Execute()
        {
            string docName = base.GetDocName("NQ");
            string fileName = Path.Combine(GetReportPath(), docName);
            base.OnShowProgress(ReportStatus.RS_START, 0);
            DataTable newQuotInfo = this.GetNewQuotInfo();
            if (newQuotInfo == null)
            {
                return null;
            }
            base.OnShowProgress(ReportStatus.RS_COMPLETEMAINDOC, 20);
            DataTable newQuotItms = this.GetNewQuotItms();
            if (newQuotItms == null)
            {
                return null;
            }
            base.OnShowProgress(ReportStatus.RS_COMPLETEDETAIL, 40);
            string path = Path.Combine(base.EagleDocPath, base.Template);
            if (!File.Exists(path))
            {
                MessageBox.Show("Cannot find the source document: " + base.Template);
                return string.Empty;
            }
            base.OnShowProgress(ReportStatus.RS_READTOMERGEDATA, 50);
            Document doc = new Document(path);
            this.mBuilder = new DocumentBuilder(doc);
            doc.MailMerge.MergeImageField += new MergeImageFieldEventHandler(this.MergeImageField);
            doc.MailMerge.MergeField += new MergeFieldEventHandler(this.MergeField);
            doc.MailMerge.Execute(newQuotInfo);
            doc.MailMerge.ExecuteWithRegions(newQuotItms);
            doc.MailMerge.DeleteFields();
            doc.Save(fileName);
            base.OnShowProgress(ReportStatus.RS_WAITFORWORDTOSTART, 60);
            return docName;
        }

        private DataTable GetNewQuotInfo()
        {
            string commandText = "select * from newquot where quot_no =" + base.RecNo;
            DataTable table = base.ExecuteDataTable(commandText);
            if ((table == null) || (table.Rows.Count == 0))
            {
                return null;
            }
            DataRow row = table.Rows[0];
            row[table.Columns.Add("quot_date_en")] = base.GetEnglishDateTime(row["quot_date"].ToString(), "MMMM d, yyyy");
            this.mCbm = row["MesUnit"].ToString().ToUpper() != "CUBE";
            this.mRatio = double.Parse(row["ratio"].ToString());
            this.mQuot_no = row["quot_no"].ToString();
            return table;
        }

        private DataTable GetNewQuotItms()
        {
            string commandText = "Select * from newquotitms Where quot_no = " + this.mQuot_no;
            DataTable table = base.ExecuteDataTable(commandText);
            if ((table == null) || (table.Rows.Count == 0))
            {
                MessageBox.Show("\x00ca\x00fd\x00be\x00dd\x00cf\x00ea\x00c7\x00e9\x00ce\x00aa\x00bf\x00d5\x00a3\x00a1");
                return null;
            }
            table.TableName = "NewQuotItms";
            DataColumn column = table.Columns.Add("mesval");
            DataColumn column2 = table.Columns.Add("itemimage");
            DataColumn column3 = table.Columns["price"];
            DataColumn column4 = table.Columns["mes"];
            foreach (DataRow row in table.Rows)
            {
                double result = 0.0;
                double.TryParse(row[column3].ToString(), out result);
                row[column3] = result * this.mRatio;
                if (this.mCbm)
                {
                    row[column] = row[column4];
                }
                else
                {
                    double num2 = 0.0;
                    double.TryParse(row[column4].ToString(), out num2);
                    row[column] = (num2 * 0x24).ToString("#0.00");
                }
                row[column2] = row["item"] as string;
            }
            return table;
        }

        private void MergeField(object sender, MergeFieldEventArgs e)
        {
        }

        private void MergeImageField(object sender, MergeImageFieldEventArgs e)
        {
            if ((e.FieldName == "itemimage") && (e.FieldValue != null))
            {
                this.mBuilder.MoveToMergeField("itemimage");
                string path = Path.Combine(base.EagleNewImagePath, e.FieldValue + ".jpg");
                if (base.IncludeImage && File.Exists(path))
                {
                    Image image = Image.FromFile(path);
                    Size adjustedSize = base.GetAdjustedSize(image, 120);
                    this.mBuilder.InsertImage(image, (double)adjustedSize.Width, (double)adjustedSize.Height);
                }
            }
        }
    }
}
