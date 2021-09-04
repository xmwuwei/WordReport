namespace WordReport
{
    using Aspose.Words;
    using Aspose.Words.Reporting;
    using System;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;

    internal class ContractReport : EagleReport
    {
        private DocumentBuilder mBuilder;
        private string mSale_no;
        private double mRareAmount;
        private string mCurr;
        private string mDis_re;

        public ContractReport(string strTemplate, bool bIncludeImage, string strRecNo, string strEaglePath) : base(strTemplate, bIncludeImage, strRecNo, strEaglePath)
        {
        }

        public override string Execute()
        {
            string docName = base.GetDocName("SC");
            string fileName = Path.Combine(GetReportPath(), docName);
            base.OnShowProgress(ReportStatus.RS_START, 0);
            DataTable contractInfo = this.GetContractInfo();
            if (contractInfo == null)
            {
                return null;
            }
            base.OnShowProgress(ReportStatus.RS_COMPLETEMAINDOC, 20);
            DataTable saleItms = this.GetSaleItms();
            if (saleItms == null)
            {
                return null;
            }
            DataRow row = contractInfo.Rows[0];
            float result = 0f;
            float.TryParse(row["discount"].ToString(), out result);
            row[contractInfo.Columns.Add("TotalAmount")] = this.mRareAmount + result;
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
            doc.MailMerge.Execute(contractInfo);
            doc.MailMerge.ExecuteWithRegions(saleItms);
            doc.MailMerge.DeleteFields();
            doc.Save(fileName);
            base.OnShowProgress(ReportStatus.RS_WAITFORWORDTOSTART, 60);
            return docName;
        }

        private DataTable GetContractInfo()
        {
            string commandText = "select * from salecon where ref_no ='" + base.RecNo + "'";
            DataTable table = base.ExecuteDataTable(commandText);
            if ((table == null) || (table.Rows.Count == 0))
            {
                return null;
            }
            DataRow row = table.Rows[0];
            row[table.Columns.Add("sale_date_en")] = base.GetEnglishDateTime(row["sale_date"].ToString(), "MMMM d, yyyy");
            this.mCurr = row["curr"].ToString();
            this.mSale_no = row["sale_no"].ToString();
            this.mDis_re = row["dis_re"].ToString();
            return table;
        }

        private DataTable GetSaleItms()
        {
            string commandText = "Select * from saleitms Where sale_no = " + this.mSale_no;
            DataTable table = base.ExecuteDataTable(commandText);
            if ((table == null) || (table.Rows.Count == 0))
            {
                MessageBox.Show("\x00ca\x00fd\x00be\x00dd\x00cf\x00ea\x00c7\x00e9\x00ce\x00aa\x00bf\x00d5\x00a3\x00a1");
                return null;
            }
            table.TableName = "SaleItms";
            DataColumn column = table.Columns.Add("amount");
            DataColumn column2 = table.Columns.Add("itemimage");
            this.mRareAmount = 0.0;
            foreach (DataRow row in table.Rows)
            {
                double num;
                double num2;
                if (!double.TryParse(row["quantity"].ToString(), out num))
                {
                    num = 0.0;
                }
                if (!double.TryParse(row["price"].ToString(), out num2))
                {
                    num2 = 0.0;
                }
                double num3 = float.Parse((num * num2).ToString("f2"));
                row[column] = num3;
                row[column2] = row["item"] as string;
                this.mRareAmount += num3;
            }
            return table;
        }

        private void MergeField(object sender, MergeFieldEventArgs e)
        {
            if ((e.FieldName.ToLower() == "discount") && (e.FieldValue != null))
            {
                float num = float.Parse(e.FieldValue.ToString());
                if (num != 0f)
                {
                    this.mBuilder.MoveToMergeField("discount");
                    this.mBuilder.Write("(" + this.mCurr + ") " + this.mRareAmount.ToString("#,##0.00") + "\n\t");
                    this.mBuilder.Write(this.mDis_re + " ");
                    this.mBuilder.Write(num.ToString("#,##0.00"));
                }
                else
                {
                    this.mBuilder.MoveToMergeField("discount");
                    this.mBuilder.Write("");
                }
            }
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
                    Size adjustedSize = base.GetAdjustedSize(image, 200);
                    this.mBuilder.InsertImage(image, (double)adjustedSize.Width, (double)adjustedSize.Height);
                }
            }
        }
    }
}
