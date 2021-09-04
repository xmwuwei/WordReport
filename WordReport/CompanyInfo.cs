namespace WordReport
{
    using System;
    using System.Xml;

    internal class CCompanyInfo
    {
        private string strTagCompanyInfo = "CompanyInfo";
        private string strTagChineseCompanyName = "ChineseCompanyName";
        private string strTagEnglishCompanyName = "EnglishCompanyName";
        private string strTagChineseCompanyAddress = "ChineseCompanyAddress";
        private string strTagEnglishCompanyAddress = "EnglishCompanyAddress";
        private string strTagTelephoneNumber = "TelephoneNumber";
        private string strTagFaxNumber = "FaxNumber";
        private string strTagEMail = "EMail";
        private string m_strChineseCompanyName;
        private string m_strEnglishCompanyName;
        private string m_strChineseCompanyAddress;
        private string m_strEnglishCompanyAddress;
        private string m_strTelephoneNumber;
        private string m_strFaxNumber;
        private string m_strEMail;

        public bool SetValueByXMLNode(XmlNode nodeCompany)
        {
            if (nodeCompany == null)
            {
                return false;
            }
            if (nodeCompany.Name != this.strTagCompanyInfo)
            {
                return false;
            }
            foreach (XmlNode node in nodeCompany.ChildNodes)
            {
                if (node.Name == this.strTagChineseCompanyName)
                {
                    this.m_strChineseCompanyName = node.InnerText;
                    continue;
                }
                if (node.Name == this.strTagEnglishCompanyName)
                {
                    this.m_strEnglishCompanyName = node.InnerText;
                    continue;
                }
                if (node.Name == this.strTagChineseCompanyAddress)
                {
                    this.m_strChineseCompanyAddress = node.InnerText;
                    continue;
                }
                if (node.Name == this.strTagEnglishCompanyAddress)
                {
                    this.m_strEnglishCompanyAddress = node.InnerText;
                    continue;
                }
                if (node.Name == this.strTagTelephoneNumber)
                {
                    this.m_strTelephoneNumber = node.InnerText;
                    continue;
                }
                if (node.Name == this.strTagFaxNumber)
                {
                    this.m_strFaxNumber = node.InnerText;
                    continue;
                }
                if (node.Name == this.strTagEMail)
                {
                    this.m_strEMail = node.InnerText;
                }
            }
            return true;
        }

        public string ChineseCompanyName
        {
            get =>
                this.m_strChineseCompanyName;
            set =>
                this.m_strChineseCompanyName = value;
        }

        public string EnglishCompanyName
        {
            get =>
                this.m_strEnglishCompanyName;
            set =>
                this.m_strEnglishCompanyName = value;
        }

        public string ChineseCompanyAddress
        {
            get =>
                this.m_strChineseCompanyAddress;
            set =>
                this.m_strChineseCompanyAddress = value;
        }

        public string EnglishCompanyAddress
        {
            get =>
                this.m_strEnglishCompanyAddress;
            set =>
                this.m_strEnglishCompanyAddress = value;
        }

        public string TelephoneNumber
        {
            get =>
                this.m_strTelephoneNumber;
            set =>
                this.m_strTelephoneNumber = value;
        }

        public string FaxNumber
        {
            get =>
                this.m_strFaxNumber;
            set =>
                this.m_strFaxNumber = value;
        }

        public string EMail
        {
            get =>
                this.m_strEMail;
            set =>
                this.m_strEMail = value;
        }
    }
}
