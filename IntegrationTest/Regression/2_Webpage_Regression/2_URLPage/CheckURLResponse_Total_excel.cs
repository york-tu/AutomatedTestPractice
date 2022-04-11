using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;
using System.Net;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class C_c_CheckURLResponse_Total_excel:IntegrationTestBase
    {
        public C_c_CheckURLResponse_Total_excel(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        #region �u�@��������}����
        /*
        �u�@�� 1:      "/bank/personal",
        �u�@�� 2:      "/bank/personal/deposit",
        �u�@�� 3:      "/bank/personal/loan", 
        �u�@�� 4:      "/bank/personal/credit-card",
        �u�@�� 5:      "/bank/personal/wealth", 
        �u�@�� 6:      "/bank/personal/trust", 
        �u�@�� 7:      "/bank/personal/insurance",
        �u�@�� 8:      "/bank/personal/lifefin", 
        �u�@�� 9:      "/bank/personal/apply", 
        �u�@�� 10:     "/bank/personal/event",
        �u�@�� 11:     "/bank/small-business", 
        �u�@�� 12:     "/bank/corporate", 
        �u�@�� 13:     "/bank/digital", 
        �u�@�� 14:     "/bank/about", 
        �u�@�� 15:     "/bank/marketing",
        �u�@�� 16:     "/bank/iframe/widget", 
        �u�@�� 17:     "/bank/error", 
        �u�@�� 18:     "/bank/bank-en",
        �u�@�� 19:     "/bank/preview";
        */
        #endregion
        [Theory]
        [InlineData (1,19)] // �X�ݤu�@��1-19������URL
         public void �ˬdURL�O�_�i�X��(int startSheetIndex, int endSheetIndex)
        {
            CreateReport("URLResponseReport", "York");

            Excel.Application excel_App = new Excel.Application(); //  new �@��excel���ε{��
            Excel.Workbook excel_WB = excel_App.Workbooks.Open(UserDataList.sitemapsExcelPath); // pass���w���|excel���eto "excel_WB"

            for (int index = startSheetIndex; index <= endSheetIndex; index++)
            {
                Excel.Worksheet excel_WS = (Excel.Worksheet)excel_WB.Worksheets[index]; //  Ū���wsheet
                INFO(excel_WS.Name);

                Excel.Range range = excel_WS.UsedRange; // ���Xsheet���e, pass to "range"
                for (int i = 0; i < range.Count; i++)
                {
                    string strURL = (string)((Excel.Range)range.Cells[i + 1, 1]).Value; // Ū�X�� i ��URL
                    try
                    {
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL); //��https request ��server
                        request.AllowAutoRedirect = false; // �����\redirect
                        request.Timeout = 30000; //�W�ɮɶ�10��

                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            PASS($"URL_{i + 1}: {strURL}    ==>    {response.StatusCode}");
                            response.Close();
                        }
                        else
                        {
                            FAIL($"URL_{i + 1}: {strURL}    ==>   {response.StatusCode}");
                            response.Close();
                        }
                    }
                    catch (WebException e)
                    {
                        WARNING($"URL_{i + 1}: {strURL}    ==>    [ Exception] {e.Message}");
                    }
                }
                INFO("");
            }

            //������������
            excel_WB.Close();
            excel_App.Quit();

            CloseBrowser(); 
        }
    }
}
