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
    public class �ˬdURL����
    {
        [Fact]

        //    "/bank/personal", "/bank/personal/deposit", "/bank/personal/loan", "/bank/personal/credit-card", "/bank/personal/wealth", 
        //    "/bank/personal/trust", "/bank/personal/insurance", "/bank/personal/lifefin", "/bank/personal/apply", "/bank/personal/event",
        //    "/bank/small-business", "/bank/corporate", "/bank/digital", "/bank/about", "/bank/marketing",
        //    "/bank/iframe/widget", "/bank/error", "/bank/bank-en", "/bank/preview" };

        //[Theory]
        //[InlineData (1,"")]
        //[InlineData(2, "")]
        //[InlineData(3, "")]
        // public void Ū��Excel��URL���ˬd�}�Ҫ�����(int caseNum, string expectResult)

        public void Ū��Excel��URL���ˬd�}�Ҫ�����()
        {




            Excel.Application excel_App = new Excel.Application(); //  new �@��excel���ε{��
            Excel.Workbook excel_WB = excel_App.Workbooks.Open(UserDataList.sitemapsExcelPath); // pass���w���|excel���eto "excel_WB"
            Excel.Worksheet excel_WS= (Excel.Worksheet)excel_WB.Worksheets[2]; //  Ū�S�wsheet
            Excel.Range range = excel_WS.UsedRange; // ���Xsheet���e, pass to "range"
            for (int i = 0; i < range.Count; i++)
            {
                var aaa = ((Excel.Range) range.Cells[i + 1, 1]).Value; // Ū�X�� i ��URL
            }


         
            int sheetNum = 1; //  ��sheetNum�Ӥu�@��



            excel_App.DisplayAlerts = false; // �s�ɮɤ����O�_�л\alert
 
            //������������
            excel_WB.Close();
            excel_App.Quit();

        }
    }
}
