namespace WordReport
{
    using Aspose.Words;
    using Aspose.Words.Reporting;
    using System;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;

    internal class PurchaseReport : EagleReport
    {
        private DocumentBuilder mBuilder;
        private string mPur_no;
        private double mRareAmount;
        private string mClerkID;

        public PurchaseReport(string strTemplate, bool bIncludeImage, string strRecNo, string strEaglePath) : base(strTemplate, bIncludeImage, strRecNo, strEaglePath)
        {
            this.mClerkID = string.Empty;
        }

        public override string Execute()
        {
            string docName = base.GetDocName("PO");
            string fileName = Path.Combine(GetReportPath(), docName);
            base.OnShowProgress(ReportStatus.RS_START, 0);
            DataTable purchaseInfo = this.GetPurchaseInfo();
            if (purchaseInfo == null)
            {
                return null;
            }
            base.OnShowProgress(ReportStatus.RS_COMPLETEMAINDOC, 20);
            DataTable purItms = this.GetPurItms();
            if (purItms == null)
            {
                return null;
            }
            DataColumn column2 = purchaseInfo.Columns.Add("ChinaTxt");
            DataColumn column3 = purchaseInfo.Columns.Add("Username");
            DataRow row = purchaseInfo.Rows[0];
            base.OnShowProgress(ReportStatus.RS_COMPLETEDETAIL, 40);
            row[purchaseInfo.Columns.Add("amount")] = this.mRareAmount;
            row[column2] = (row["chinalabel"].ToString() != "YES") ? "每件商品不需有MADE IN CHINA 标志" : "每件商品需有MADE IN CHINA 标志";
            row[column3] = this.GetUserName();
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
            doc.MailMerge.Execute(purchaseInfo);
            doc.MailMerge.ExecuteWithRegions(purItms);
            doc.MailMerge.DeleteFields();
            doc.Save(fileName);
            base.OnShowProgress(ReportStatus.RS_WAITFORWORDTOSTART, 60);
            return docName;
        }

        private DataTable GetPurchaseInfo()
        {
            string commandText = "select * from pur where ref_no ='" + base.RecNo + "'";
            DataTable table = base.ExecuteDataTable(commandText);
            if ((table == null) || (table.Rows.Count == 0))
            {
                return null;
            }
            DataRow row = table.Rows[0];
            this.mPur_no = row["pur_no"].ToString();
            this.mClerkID = row["ClerkID"].ToString();
            return table;
        }

        private DataTable GetPurItms()
        {
            string commandText = "Select * from puritms Where pur_no = " + this.mPur_no;
            DataTable table = base.ExecuteDataTable(commandText);
            if ((table == null) || (table.Rows.Count == 0))
            {
                MessageBox.Show("\x00ca\x00fd\x00be\x00dd\x00cf\x00ea\x00c7\x00e9\x00ce\x00aa\x00bf\x00d5\x00a3\x00a1");
                return null;
            }
            table.TableName = "PurItms";
            DataColumn column = table.Columns.Add("itemimage");
            DataColumn column2 = table.Columns.Add("ctns");
            this.mRareAmount = 0.0;
            foreach (DataRow row in table.Rows)
            {
                double num;
                double num2;
                double num3;
                if (!double.TryParse(row["quantity"].ToString(), out num))
                {
                    num = 0.0;
                }
                if (!double.TryParse(row["cost"].ToString(), out num2))
                {
                    num2 = 0.0;
                }
                if (!double.TryParse(row["packing"].ToString(), out num3))
                {
                    num3 = 1.0;
                }
                if (Math.Abs(num3) < 1E-06)
                {
                    num3 = 1.0;
                }
                double num4 = double.Parse((num * num2).ToString("f2"));
                row[column] = row["item"] as string;
                row[column2] = ((double)((int)num)) / num3;
                this.mRareAmount += num4;
            }
            return table;
        }

        private string GetUserName()
        {
            if (string.IsNullOrEmpty(this.mClerkID))
            {
                return string.Empty;
            }
            string commandText = "Select Username from Clerk Where ClerkID = " + this.mClerkID;
            DataTable table = base.ExecuteDataTable(commandText);
            return ((table.Rows.Count != 0) ? table.Rows[0]["Username"].ToString() : string.Empty);
        }

        private void MergeField(object sender, MergeFieldEventArgs e)
        {
        }

        private void MergeImageField(object sender, MergeImageFieldEventArgs e)
        {
            if ((e.FieldName == "itemimage") && (e.FieldValue != null))
            {
                this.mBuilder.MoveToMergeField("itemimage");
                string path = Path.Combine(base.EagleImagePath, e.FieldValue + ".jpg");
                if (base.IncludeImage && File.Exists(path))
                {
                    Image image = Image.FromFile(path);
                    Size adjustedSize = base.GetAdjustedSize(image, 140);
                    this.mBuilder.InsertImage(image, (double)adjustedSize.Width, (double)adjustedSize.Height);
                    this.mBuilder.Writeln();
                }
            }
        }
    }
}
