using Xunit;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using AutomatedTest.Utilities;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class A_QueryURL
    {
        [Fact]
        public void ����URL�öץX��excel()
        {
            string excelExportPath = $@"{UserDataList.Upperfolderpath}xmlURLList";

            #region ��������l�X > �s�Jstring "sitemaps" > �r�갵���� > �s�J list<>totalURLList
            HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("https://www.esunbank.com.tw/bank/sitemap.xml");
            request1.Method = "GET";
            request1.AllowAutoRedirect = true;
            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();

            StreamReader reader = new StreamReader(response1.GetResponseStream());
            string sitemaps = reader.ReadToEnd(); //���X������l�X�æs�Jstring "sitemaps"
            List<string> totalURLList = sitemaps.Split('\n').ToList();
            #endregion

            List<string> urlList = new List<string>();
            string pattern = @"((http|ftp|https)://)(([a-zA-Z0-9\._-]+\.[a-zA-Z]{2,6})|([0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}))(:[0-9]{1,4})*(/[a-zA-Z0-9\&%_\./-~-]*)?"; // �P�_���}URL
            foreach (var urlLink in totalURLList)
            {
                MatchCollection matchSources = Regex.Matches(urlLink, pattern); // ��X�Ҧ�URL�r��
                for (int i = 0; i < matchSources.Count; i++)
                {
                    if (matchSources[i].ToString().Contains("esunbank")) // �P�_�r��O�_�]�tesunbank
                    {
                        urlList.Add(matchSources[i].ToString().Replace("</loc>", "")); // URL�g�JurlList
                    }
                }
            }
            Excel.Application excel_App1 = new Excel.Application(); //  new �@��excel���ε{��
            Excel.Workbook excel_WB1 = excel_App1.Workbooks.Add(); // �s�Wexcel�ɮ�
            Excel.Worksheet excel_WS1 = new Excel.Worksheet(); //  �s�W�u�@��

            string[] structureList = new string[] { "/bank/personal", "/bank/personal/deposit", "/bank/personal/loan", "/bank/personal/credit-card", "/bank/personal/wealth",
                "/bank/personal/trust", "/bank/personal/insurance", "/bank/personal/lifefin", "/bank/personal/apply", "/bank/personal/event",
                "/bank/small-business", "/bank/corporate", "/bank/digital", "/bank/about", "/bank/marketing",
                "/bank/iframe/widget", "/bank/error", "/bank/bank-en", "/bank/preview" };

            excel_WB1.Worksheets.Add(Missing.Value, Missing.Value, structureList.Length - 1, Missing.Value); // �s�W�u�@��ƶq (total�ƶq = structureList�Ӽ� -1 + �w�]1)

            int sheetNum = 1; //  ��sheetNum�Ӥu�@��
            foreach (var item in structureList)
            {
                excel_WS1 = (Excel.Worksheet)excel_WB1.Worksheets[sheetNum]; // �����sheetNum�Ӥu�@��
                excel_WS1.Name = item.Replace('/', '_').Substring(6); //�R�W�u�@��
                excel_WS1.Activate(); // setup�u�@��J�I

                int index = 0; // excel �}�l���
                for (int i = 0; i < urlList.Count; i++)
                {
                    if (item == "/bank/personal" && urlList[i].EndsWith("bank/personal")) // for�u�@��"personal"���e
                    {
                        excel_App1.Cells[index + 1, 1] = urlList[i]; // �g�Jexcel
                        index++;
                        break;
                    }
                    if (urlList[i].Contains(item)) // �P�_URL�O�_�tstructureList���r��
                    {
                        excel_App1.Cells[index + 1, 1] = urlList[i]; // �g�Jexcel
                        index++;
                    }
                }
                sheetNum++;
            }
            ((Excel.Worksheet)excel_WB1.Worksheets[1]).Activate(); // �]�w�u�@��focus on �u�@��1

            excel_App1.DisplayAlerts = false; // �s�ɮɤ����O�_�л\alert
            excel_WB1.SaveAs(excelExportPath); // �s��

            //������������
            excel_WB1.Close();
            excel_App1.Quit();

        }
    }
}
