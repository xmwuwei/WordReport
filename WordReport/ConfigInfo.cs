namespace WordReport
{
    using System;
    using System.Xml;

    internal class CConfigInfo
    {
        private const int nCompanyCount = 2;
        private static CConfigInfo instance;
        private string strTagCompanyInfos = "CompanyInfos";
        private string strTagCompanyInfo = "CompanyInfo";
        private CCompanyInfo[] m_Company = new CCompanyInfo[2];

        private CConfigInfo()
        {
            for (int i = 0; i < 2; i++)
            {
                this.m_Company[i] = new CCompanyInfo();
            }
        }

        public bool GetInfoFromXML()
        {
            XmlDocument document = new XmlDocument();
            document.Load(@"C:\Config.xml");
            foreach (XmlNode node in document.DocumentElement.ChildNodes)
            {
                if (node.Name == this.strTagCompanyInfos)
                {
                    XmlNodeList list = node.SelectNodes(this.strTagCompanyInfo);
                    for (int i = 0; i < Math.Min(list.Count, 2); i++)
                    {
                        XmlNode nodeCompany = list[i];
                        if (nodeCompany != null)
                        {
                            this.m_Company[i].SetValueByXMLNode(nodeCompany);
                        }
                    }
                }
            }
            return true;
        }

        public static CConfigInfo Instance
        {
            get
            {
                instance = new CConfigInfo();
                return instance;
            }
        }
    }
}
