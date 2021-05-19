using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;

namespace 預售屋信託查詢
{
    public class 預售屋信託查詢
    {
        readonly string test_url = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        readonly string excel_path = @"D:\PreSaleHouseTrustInquiry_TestReport.xls";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void PreSaleHouseTrustInquiry(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);

            using var driver = WebDriverInfra.Create_Browser(browserType);
            {
                string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");

                WorkBook xlsWorkbook;
                WorkSheet xlsSheet;
                if (File.Exists(excel_path) != true)
                {
                    xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLS); //定義 excel格式為 XLS
                    xlsSheet = xlsWorkbook.CreateWorkSheet("預售屋查詢" + timesavepath); // Create excel檔 "sheet"分頁
                }
                else
                {
                    xlsWorkbook = WorkBook.Load(excel_path); //讀取 excel 檔
                    xlsSheet = xlsWorkbook.CreateWorkSheet("預售屋查詢" + timesavepath); // Create excel檔 "sheet"分頁
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
                string[] array_payment_account = new string[] {"", "1234567890123456789","12345","１２３４５６", "中文字", "ＡＢＣＤＥＦ", "ａｂｃｄｅｆ", "abcdef", "ABCDEF", "@#$%^&*", "！＠＃＄％︿", "123abc456DEF", "987654321"};
                
                //定義excel欄位
                xlsSheet["A1"].Value = "檢查 繳款帳號欄位";
                xlsSheet["A2"].Value = "建案名稱";
                xlsSheet["B2"].Value = "手動輸入";
                xlsSheet["C2"].Value = "欄位實際值";
                xlsSheet["D2"].Value = "顯示訊息";
                xlsSheet["E2"].Value = "預期結果";
                xlsSheet["F2"].Value = "測試結果";
                xlsSheet["A2:F2"].Style.BottomBorder.SetColor("#ff6600"); // 底線"紅色"
                xlsSheet["A2:F2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double; //加雙底線
                xlsSheet["A3"].Value = ProjectNameDropDownList.Text; // 欄位=建案名稱


                int i = 3;
                foreach (var keyin in array_payment_account) // 跑迴圈以逐筆輸入欄位
                {
                    string check_position = "B" + i;
                    string actual_value = "C" + i;
                    string show_msg = "D" + i;
                    string expect_result = "E" + i;
                    string test_result = "F" + i;
                    

                    bool result = Regex.IsMatch(keyin, @"^[0-9]*$"); // 判斷輸入字串是否為"全數字"

                    IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
                    PaymentAccount.Clear(); // 清除繳款欄位
                    PaymentAccount.SendKeys(keyin); // 輸入值

                    string actualPaymentAccount = PaymentAccount.GetAttribute("value");

                    IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));
                   
                    if (keyin.Length >= 6 && result == true) // 判斷當輸入字元為全數字且大於六位數時, 送出data後延遲等待5秒, 等網頁load完
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

                    if (keyin == "" && payment_account_error == "必須填寫")// 判斷當輸入為"空值" >>> 預期顯示 "必須填寫" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '必須填寫'";
                        xlsSheet[test_result].Value = "PASS_Case_1";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_1: " + keyin);
                    }
                    else if (keyin.Length < 6 && payment_account_error == "最少 6 個字") // 判斷當輸入字元"小於六位數" >>> 預期顯示 "最少 6 個字" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '最少 6 個字'";
                        xlsSheet[test_result].Value = "PASS_Case_2";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_2: " + keyin);
                    }
                    else if (result != true && payment_account_error == "只可輸入數字") // 判斷當輸入"不是全數字" >>> 預期顯示 "只可輸入數字" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '只可輸入數字'";
                        xlsSheet[test_result].Value = "PASS_Case_3";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_3: " + keyin);
                    }

                    // 判斷當輸入 "全部都數字" 且 "字數介於 6~16 位數間 >>> 預期查詢結果顯示 "企業代碼不存在"
                    else if (result == true && keyin.Length >=6 && keyin.Length <=16 && driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text == "企業代碼不存在")
                    {   
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = searchresult;
                        xlsSheet[expect_result].Value = "顯示 '企業代碼不存在'";
                        xlsSheet[test_result].Value = "PASS_Case_4";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_4: " + keyin);
                    }
                    else if (keyin.Length > 16 && result == true) // 判斷當輸入位元 "大於16位數" 且 "全部都數字" >> 預期查詢結果顯示 "企業代碼不存在 (need to confirm)"
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = searchresult;
                        xlsSheet[expect_result].Value = "顯示 '[TBD]企業代碼不存在'";
                        xlsSheet[test_result].Value = "Need Manaul Check";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_5: " + keyin);
                    }
                    else //非以上情況測試結果fail
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = "Message 1: " + payment_account_error + "or Message 2: " + searchresult;
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
