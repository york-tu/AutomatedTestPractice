using OpenQA.Selenium;
using Xunit;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System;

namespace AutomatedTest.IntegrationTest.Personal.Trust
{
    public class 預售屋信託查詢精進_PC版: IntegrationTestBase
    {
        public 預售屋信託查詢精進_PC版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        }

        private readonly string testcase_name = "預售屋查詢";

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 欄位檢核測試(string browser)
        {
            StartTestCase(browser, "預售屋信託查詢精進_欄位檢核測試", "York");
            INFO("");
            string time = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            string snapshotfolderpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\PreSaleHouseTrustInquiry";
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

                IWebElement ProjectNameDropDownList = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul"));
                ProjectNameDropDownList.Click(); // 點建案下拉選單
                System.Threading.Thread.Sleep(100);
                IWebElement SelectProject = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul/li/ul/li[6]/span"));
                SelectProject.Click();// 選建案
                System.Threading.Thread.Sleep(100);
                
                IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));
                SubmitButton.Click();

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

                    

                    string actualPaymentAccountColumn = PaymentAccount.GetAttribute("value"); // 實際欄位讀取到的值

                    bool regexCheckResult = Regex.IsMatch(actualPaymentAccountColumn, @"^\d{6,16}$"); // 判斷欄位裡的字串是否為"6-16位全數字"
                    bool numericResult = Regex.IsMatch(actualPaymentAccountColumn, @"^[0-9]*$"); // 判斷欄位裡的字串是否為"全數字"
                
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // 欄位錯誤訊息

                    if (regexCheckResult == true) // 判斷當輸入字元符合預期(全數字, 6~16位數), 送出data後延遲等待5秒, 等網頁load完
                    {
                    SubmitButton.Click(); // 點 送出
                        System.Threading.Thread.Sleep(5000);
                    }
                    else
                    {
                    SubmitButton.Click(); // 點 送出
                        System.Threading.Thread.Sleep(100);
                    }


                    if (string.IsNullOrWhiteSpace(keyin) == true)// 判斷當輸入欄位為"空值" >>> 預期顯示 "必須填寫" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '必須填寫'";
                        xlsSheet[test_result].Value = "Case_1";
                        try
                        {
                           Assert.Equal("必須填寫", payment_account_error);
                            PASS($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 必須填寫");
                        }
                         catch (System.Exception e)
                        {
                           ERROR(e.Message);
                           FAIL($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 必須填寫");
                           Assert.True(false);
                        }
                    }


                    else if (string.IsNullOrWhiteSpace(actualPaymentAccountColumn) != true && numericResult != true) //  判斷當輸入非數字  >>> 預期顯示 "只可輸入數字" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '只可輸入數字'";
                        xlsSheet[test_result].Value = "Case_3";
                       try
                       {
                            Assert.Equal("只可輸入數字", payment_account_error);
                            PASS($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 只可輸入數字");
                       }
                       catch (System.Exception e)
                       {
                           ERROR(e.Message);
                           FAIL($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 只可輸入數字");
                           Assert.True(false);
                       }
                    }
                    else if (string.IsNullOrWhiteSpace(actualPaymentAccountColumn) != true && regexCheckResult != true) // 判斷當"輸入欄位數字小於六位數" >>> 預期顯示 "最少 6 個字" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "顯示 '最少 6 個字'";
                        xlsSheet[test_result].Value = "Case_2";

                         try
                         {
                            Assert.Equal("最少6個字", payment_account_error);
                            PASS($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 最少6個字");
                         }
                        catch (System.Exception e)
                         {
                           ERROR(e.Message);
                           FAIL($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 最少6個字");
                           Assert.True(false);
                         }
                    }
                    else if (string.IsNullOrWhiteSpace(actualPaymentAccountColumn) != true && regexCheckResult == true) // 輸入 "全部都數字" 且 "字數介於 6~16 位數間
                    {
                    retry:
                        SubmitButton.Click();
                        System.Threading.Thread.Sleep(15000);
                        IWebElement verifycode = driver.FindElement(By.XPath("")); // 驗證碼欄位
                        verifycode.Clear();
                        verifycode.Click();
                        System.Threading.Thread.Sleep(15000);

                        SubmitButton.Click();

                        System.Threading.Thread.Sleep(15000);
                        string verifycode_error = driver.FindElement(By.Id("captchaWrong")).Text; // 驗證碼錯誤訊息
                        string actualVerifyCodeColumnValue = verifycode.GetAttribute("value"); // 驗證碼欄位實際讀到的值
                          if (string.IsNullOrWhiteSpace(actualVerifyCodeColumnValue) == true) // 輸入空驗證碼
                          {
                             try
                             {
                                Assert.Equal("請輸入驗證碼", payment_account_error);
                                PASS($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 請輸入驗證碼");
                             }
                            catch (System.Exception e)
                             {
                               ERROR(e.Message);
                               FAIL($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 請輸入驗證碼");
                               Assert.True(false);
                             }
                          goto retry;
                         }
                          else if (string.IsNullOrWhiteSpace(actualVerifyCodeColumnValue) != true && string.IsNullOrWhiteSpace(verifycode_error) != true)
                          {
                            try
                             {
                                Assert.Equal("驗證碼錯誤", payment_account_error);
                                PASS($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 驗證碼錯誤");
                             }
                            catch (System.Exception e)
                             {
                               ERROR(e.Message);
                               FAIL($"輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 驗證碼錯誤");
                               Assert.True(false);
                             }
                            goto retry;
                          }
                          else if (string.IsNullOrWhiteSpace(actualVerifyCodeColumnValue) != true && string.IsNullOrWhiteSpace(verifycode_error) == true)
                           {
                                
                                INFO("");
                                INFO("檢查查詢結果");

                                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                                
                                if (string.IsNullOrWhiteSpace(searchresult) != true)
                                {
                                xlsSheet[check_position].Value = keyin;
                                xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                                xlsSheet[show_msg].Value = searchresult;
                                xlsSheet[expect_result].Value = "顯示 '企業代碼不存在 / 繳款帳號不存在'";
                                xlsSheet[test_result].Value = "Case_4";

                                WARNING($"[NeedCheck][正確驗證碼], 輸入帳號: {keyin}, 輸入長度{keyin.Length}, 欄位實際讀到: {actualPaymentAccountColumn}, 讀到長度: {actualPaymentAccountColumn.Length}, 實際錯誤訊息: {payment_account_error}, 預期錯誤訊息: 企業代碼不存在 / 繳款帳號不存在");
                                }

                                bool date_column = driver.FindElement(By.XPath("")).Displayed; // Result 表單 "日期" 欄位顯示
                                bool summary_column = driver.FindElement(By.XPath("")).Displayed; // Result 表單 "摘要" 欄位顯示
                                bool expenditure_column = driver.FindElement(By.XPath("")).Displayed; // Result 表單 "支出" 欄位顯示
                                bool deposit_column = driver.FindElement(By.XPath("")).Displayed; // Result 表單 "存入" 欄位顯示
                                bool surolus_column = driver.FindElement(By.XPath("")).Displayed; // Result 表單 "結餘" 欄位顯示
                                
                                try
                                {
                                    Assert.True(date_column);
                                    Assert.True(summary_column );
                                    Assert.True(expenditure_column);
                                    Assert.True(deposit_column);
                                    Assert.True(surolus_column);
                                    PASS("查詢結果符合預期");
                                }
                                catch (Exception e)
                                {   
                                    ERROR(e.Message);
                                    PASS("查詢結果不符合預期");
                                    Assert.True(false);
                                }
                          }
                    }
                }

                if (File.Exists(excel_path) != true)
                {
                    xlsWorkbook.SaveAs(excel_path);
                }
                else
                {
                    xlsWorkbook.Save();
                }
            CloseBrowser();
        }
    }
}
