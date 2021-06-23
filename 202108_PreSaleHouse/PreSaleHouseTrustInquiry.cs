using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using System;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;
using Utilities;

namespace PreSaleHouseTrustInquiryTest
{
    public class 預售屋信託查詢精進_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        
        private readonly string testcase_name = "預售屋查詢";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);

            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                string tim = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");
                string snapshotfolderpath = $@"{UserDataList.folderpath}\SnapshotFolder\PreSaleHouseTrustInquiry";
                string excel_path = $@"{snapshotfolderpath}\TestReport.xlsx";

                Tools.CreateSnapshotFolder(snapshotfolderpath);
                System.Threading.Thread.Sleep(100);


                WorkBook xlsWorkbook;
                WorkSheet xlsSheet;

                if (File.Exists(excel_path) == true && WorkBook.Load(excel_path).GetWorkSheet(testcase_name) == null)
                { // 判斷當excel檔存在 但 沒有"預售屋查詢"工作表 >>> 讀取該excel檔, new create "預售屋查詢" 工作表 
                    xlsWorkbook = WorkBook.Load(excel_path); 
                    xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name); 
                }
                else if (File.Exists(excel_path) == true && WorkBook.Load(excel_path).GetWorkSheet(testcase_name) != null)
                { // 判斷當excel檔存在 且 有"預售屋查詢"工作表 >>> 讀取該excel檔, 讀取"預售屋查詢" 工作表
                    xlsWorkbook = WorkBook.Load(excel_path);
                    xlsSheet = xlsWorkbook.GetWorkSheet(testcase_name); 
                }
                else
                { // 判斷當excel檔不存在 >>> new create excel檔 & new create "預售屋查詢" 工作表
                    xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLSX); //定義 excel格式為 XLSX
                    xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name);
                }

                
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                //driver.Manage().Window.Maximize();

                IWebElement ProjectNameDropDownList = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul"));
                ProjectNameDropDownList.Click(); // 點建案下拉選單
                System.Threading.Thread.Sleep(100);
                IWebElement SelectProject = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul/li/ul/li[6]/span"));
                SelectProject.Click();// 選建案
                System.Threading.Thread.Sleep(100);


                //Array存取預測試的字串
                string[] array_payment_account = new string[] {"", "1234567890123456789","12345","１２３４５６", "中文字", "ＡＢＣＤＥＦ", "ａｂｃｄｅｆ", "abcdef", "ABCDEF", "@#$%^&*", "！＠＃＄％︿", "123abc456DEF", "987654321", ""};
                
                //定義excel欄位
                xlsSheet["A1"].Value = "檢查 繳款帳號欄位";
                xlsSheet["A2"].Value = "建案名稱";
                xlsSheet["B2"].Value = "手動輸入";
                xlsSheet["C2"].Value = "欄位實際值";
                xlsSheet["D2"].Value = "顯示訊息";
                xlsSheet["E2"].Value = "預期結果(顯示訊息)";
                xlsSheet["F2"].Value = "測試結果";
                xlsSheet["A2:F2"].Style.BottomBorder.SetColor("#ff6600"); // 底線"紅色"
                xlsSheet["A2:F2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double; //加雙底線
                xlsSheet["A3"].Value = ProjectNameDropDownList.Text; // 欄位=建案名稱


                int i = 3;
                foreach (string keyin in array_payment_account) // 跑迴圈以逐筆輸入欄位
                {
                    string check_position = "B" + i;
                    string actual_value = "C" + i;
                    string show_msg = "D" + i;
                    string expect_result = "E" + i;
                    string test_result = "F" + i;

                    IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
                    PaymentAccount.Clear(); // 清除繳款欄位
                    PaymentAccount.SendKeys(keyin); // 輸入值

                    IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));

                    string actualPaymentAccount = PaymentAccount.GetAttribute("value"); // 實際欄位讀取到的值

                    bool expect_check_result = Regex.IsMatch(actualPaymentAccount, @"^\d{6,16}$"); // 判斷欄位裡的字串是否為"6-16位全數字"
                    bool digital_check_result = Regex.IsMatch(actualPaymentAccount, @"^[0-9]*$"); // 判斷欄位裡的字串是否為"全數字"
                

                    if (expect_check_result == true) // 判斷當輸入字元符合預期(全數字, 6~16位數), 送出data後延遲等待5秒, 等網頁load完
                    {
                        submit_button.Click(); // 點 送出
                        System.Threading.Thread.Sleep(5000);
                    }
                    else
                    {
                        submit_button.Click(); // 點 送出
                        System.Threading.Thread.Sleep(100);
                    }

                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;

                    if (actualPaymentAccount == "")// 判斷當輸入欄位為"空值" >>> 預期顯示 "必須填寫" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '必須填寫'";
                        xlsSheet[test_result].Value = "Case_1";
                    }
                    else if (actualPaymentAccount.Length < 6 && digital_check_result == true) // 判斷當"輸入欄位數字小於六位數" >>> 預期顯示 "最少 6 個字" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '最少 6 個字'";
                        xlsSheet[test_result].Value = "Case_2";
                    }
                    else if (digital_check_result != true) // 判斷當輸入欄位"不全為數字" >>> 預期顯示 "只可輸入數字" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '只可輸入數字'";
                        xlsSheet[test_result].Value = "Case_3";
                    }

                    // 判斷當輸入 "全部都數字" 且 "字數介於 6~16 位數間 >>> 預期查詢結果顯示 "企業代碼不存在"
                    else if (expect_check_result == true)
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = searchresult;
                        xlsSheet[expect_result].Value = "顯示 '企業代碼不存在'";
                        xlsSheet[test_result].Value = "Case_4";
                    }
                    //else if (keyin.Length > 16 && result == true) // 判斷當輸入位元 "大於16位數" 且 "全部都數字" >> 預期查詢結果顯示 "企業代碼不存在 (need to confirm)"
                    //{
                    //    string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                    //    xlsSheet[check_position].Value = keyin;
                    //    xlsSheet[actual_value].Value = actualPaymentAccount;
                    //    xlsSheet[show_msg].Value = searchresult;
                    //    xlsSheet[expect_result].Value = "顯示 '[TBD]企業代碼不存在'";
                    //    xlsSheet[test_result].Value = "Need Manaul Check";
                    //}
                    else //非以上情況測試結果fail
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = $"Message 1: {payment_account_error} / Message 2: {searchresult}";
                        xlsSheet[test_result].Value = "FAIL";
                    }
                    i++;
                }

                if (File.Exists(excel_path) != true)
                {
                    xlsWorkbook.SaveAs(excel_path);
                }
                else
                {
                    xlsWorkbook.Save();
                }
            }
            driver.Quit();
        }
    }
}
