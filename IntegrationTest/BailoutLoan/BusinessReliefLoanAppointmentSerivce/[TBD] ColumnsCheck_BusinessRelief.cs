using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.Text.RegularExpressions;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.BailoutLoan.LaborReliefLoanAppointmentService
{
    public class 企業紓困貸款預約服務_欄位檢核:IntegrationTestBase
    {
        public 企業紓困貸款預約服務_欄位檢核(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/small-business/tools/apply/sbloan-appointment";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void 欄位檢核(string browser)
        {
            StartTestCase(browser, "企業紓困貸款預約服務_欄位檢核", "York");
            INFO("確認欄位檢核正確");


            INFO("檢查網頁上hyperlink是否正確");
            string business_bailout_zone_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[1]/span/a")).GetAttribute("href"); // 獲取 "企業紓困專區" 的超連結網址
            string expect_business_bailout_zone_path = "https://www.esunbank.com.tw/bank/marketing/loan/rescueplan-all#sortorder400";
            if (business_bailout_zone_hyperlink == expect_business_bailout_zone_path)
            {
                PASS("企業紓困專區 hyperlink 符合預期");
            }
            else
            {
                FAIL($"企業紓困專區 link: {business_bailout_zone_hyperlink}, 預期 link: {expect_business_bailout_zone_path}");
            }
            string existing_loan_relief_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[2]/a")).GetAttribute("href"); // 獲取 既有貸款寬緩 的 "超連結" 網址
            string expect_existing_loan_relief_path = "https://www.esunbank.com.tw/bank/small-business/tools/apply/contact-me";
            if (existing_loan_relief_hyperlink == expect_existing_loan_relief_path)
            {
                PASS("既有貸款寬緩 hyperlink 符合預期");
            }
            else
            {
                FAIL($"既有貸款寬緩 link: {existing_loan_relief_hyperlink}, 預期 link: {expect_existing_loan_relief_path}");
            }
            IWebElement ExistLoanReliefButton = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[1]")); 
            string existing_loan_relief_button = ExistLoanReliefButton.GetAttribute("href"); //獲取 既有貸款寬緩 "按鈕" 的連結網址
            string expect_existing_loan_relief_button_path = "https://www.esunbank.com.tw/bank/small-business/tools/apply/contact-me";
            if(existing_loan_relief_button == expect_existing_loan_relief_button_path)
            {
                PASS("既有貸款寬緩'按鍵'連結符合預期");
            }
            else
            {
                FAIL($"既有貸款寬緩'按鍵'連結: {existing_loan_relief_button}, 預期 link: {expect_existing_loan_relief_button_path}");
            }

            
            ///<summary>
            /// 檢查網頁 "上方文案"
            ///</summary>
            INFO("");
            INFO("檢查網頁上方文案");
            string upper_wordings_first_line = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[1]")).Text;
            string upper_wordings_second_line = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[2]")).Text;
            string expect_upper_first_line_wordings = "※因防疫期間分行實施人流控管，建議可至企業紓困專區詳閱準備資料資訊，再進行線上預約服務，以加速申辦流程。";
            string expect_upper_second_line_wordings = "※本頁面為企業紓困新貸案件專用，申請「既有貸款寬緩」，請點選既有貸款寬緩前往預約頁面。";
            if (upper_wordings_first_line == expect_upper_first_line_wordings)
            {
                PASS("上方文案區文案1: 符合預期");
            }
            else
            {
                FAIL($"上方文案區文案1敘述: {upper_wordings_first_line}, 預期 link: {expect_upper_first_line_wordings}");
            }
            if (upper_wordings_second_line == expect_upper_second_line_wordings)
            {
                PASS("上方文案區文案2: 符合預期");
            }
            else
            {
                FAIL($"上方文案區文案2敘述: {upper_wordings_second_line}, 預期 link: {expect_upper_second_line_wordings}");
            }
                    
            IWebElement TelephoneExtensionColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[3]"));
            TelephoneExtensionColumn.SendKeys("QWE!@#$%"); // 分機欄位輸入錯誤文字
            IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));
            SubmitButton.Click(); // 點 送出


     
            ///<summary>
            ///公司名稱欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 公司名稱");
            string[] CompanyNameArray = new string[] {"", "       ", "Ａｂ", "ab", "Ａ", "中", "英文ａｂｃ中文Ｄｅｆ符號#$%!@~&*^%", "中文字加英文字加數字_ＡＢＣＤＥＦＧＨＩＪ＿abcdefghij_1234567890_ABCDEFGHIJ_", "GoldCompany"};
            foreach (var companyName in CompanyNameArray)
            {
                IWebElement CompanyNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]/input"));
                CompanyNameColumn.Clear();
                CompanyNameColumn.SendKeys(companyName); // 填公司名稱
                string companyNameErrorWordings = driver.FindElement(By.Id("companyName-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnCompanyName = CompanyNameColumn.GetAttribute("value"); // 實際欄位讀到的值

                if(string.IsNullOrWhiteSpace(companyName) == true)
                {
                    if(companyNameErrorWordings == "必須填寫")
                    {
                        PASS($"輸入: {companyName}, 錯誤訊息: {companyNameErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {companyName}, 錯誤訊息: {companyNameErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else if (string.IsNullOrWhiteSpace(companyName) != true && companyName.Length < 2)
                {
                    if(companyNameErrorWordings == "最少2個字")
                    {
                        PASS($"輸入: {companyName}, 錯誤訊息: {companyNameErrorWordings}, 預期錯誤訊息: 最少2個字");
                    }
                    else
                    {
                        FAIL($"輸入: {companyName}, 錯誤訊息: {companyNameErrorWordings}, 預期錯誤訊息: 最少2個字");
                    }
                }
                else if (companyName.Length >= 50)
                {
                    if(actualColumnCompanyName.Length == 50)
                    {
                        PASS($"輸入: {companyName}, 輸入長度:{companyName.Length}, 欄位長度: {actualColumnCompanyName.Length}, 預期: 長度限制50字元");
                    }
                    else
                    {
                        FAIL($"輸入: {companyName}, 輸入長度:{companyName.Length}, 欄位長度: {actualColumnCompanyName.Length}, 預期: 長度限制50字元");
                    }
                }
                else
                {
                    WARNING($"NeedCheck, 輸入: {companyName}, 錯誤訊息: {companyNameErrorWordings}");
                }    
            }



            ///<summary>
            ///公司統編欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 公司統編");
            string[] CompanyUniformNumbersArray = new string[] {"", "       ", "12345", "ａｂｃａｂｃＤｅｆ", "abcdefgh", "中文字加英文數字", "!@#$%^&*(", "1232148123123", "55665566", "@英%^&文b字c", "91301434666666", "91301434"};
            foreach (var companyUniformNumber in CompanyUniformNumbersArray)
            {
                IWebElement CompanyUniformNumbersColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]/input"));
                CompanyUniformNumbersColumn.Clear();
                CompanyUniformNumbersColumn.SendKeys(companyUniformNumber); // 填公司統編
                string companyUniformNumbersErrorWordings = driver.FindElement(By.Id("vatNumber-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnCompanyUniformNumbers = CompanyUniformNumbersColumn.GetAttribute("value"); // 實際欄位讀到的值

                if (string.IsNullOrWhiteSpace(companyUniformNumber) == true )
                {
                    if(companyUniformNumbersErrorWordings == "必須填寫")
                    {
                        PASS($"輸入: {companyUniformNumber}, 錯誤訊息: {companyUniformNumbersErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {companyUniformNumber}, 錯誤訊息: {companyUniformNumbersErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else if (Tools.CheckUniformNumber(actualColumnCompanyUniformNumbers) != true)
                {
                    if(companyUniformNumbersErrorWordings == "統一編號錯誤")
                    {
                        PASS($"輸入: {companyUniformNumber}, 輸入長度: {companyUniformNumber.Length}, 欄位實際讀到: {actualColumnCompanyUniformNumbers}, 欄位長度: {actualColumnCompanyUniformNumbers.Length}, 錯誤訊息: {companyUniformNumbersErrorWordings}, 預期錯誤訊息: 統一編號錯誤, 預期長度: 8 字元");
                    }
                    else
                    {
                        FAIL($"輸入: {companyUniformNumber}, 輸入長度: {companyUniformNumber.Length}, 欄位實際讀到: {actualColumnCompanyUniformNumbers}, 欄位長度: {actualColumnCompanyUniformNumbers.Length}, 錯誤訊息: {companyUniformNumbersErrorWordings}, 預期錯誤訊息: 統一編號錯誤, 預期長度: 8 字元");
                    }
                }
                else if (Tools.CheckUniformNumber(actualColumnCompanyUniformNumbers) == true)
                {
                       PASS($"輸入: {companyUniformNumber}, 輸入長度: {companyUniformNumber.Length}, 欄位實際讀到: {actualColumnCompanyUniformNumbers}, 欄位長度: {actualColumnCompanyUniformNumbers.Length}, 錯誤訊息: {companyUniformNumbersErrorWordings}, 預期錯誤訊息: n/a, 預期長度: 8 字元");
                }
            }



             ///<summary>
            ///負責人欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 負責人");
            string[] PrincipalArray = new string[] {"", "       ", "Ａｂ", "ab", "許功蓋ㄟㄟ", "許功蓋Bluce", "Ａ", "中", "英文ａｂｃ中文Ｄｅｆ符號#$%!@~&*^%", "中文字加英文字加數字_ＡＢＣＤＥＦＧＨＩＪ＿abcdefghij_1234567890_ABCDEFGHIJ_", "GoldPrinciple"};
            foreach (var Principal in PrincipalArray)
            {
                IWebElement PrincipalColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[2]/input"));
                PrincipalColumn.Clear();
                PrincipalColumn.SendKeys(Principal); // 填負責人
                string principalColumnErrorWordings = driver.FindElement(By.Id("name-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnPrincipal = PrincipalColumn.GetAttribute("value"); // 實際欄位讀到的值

                if(string.IsNullOrWhiteSpace(Principal) == true)
                {
                    if(principalColumnErrorWordings == "必須填寫")
                    {
                        PASS($"輸入: {Principal}, 錯誤訊息: {principalColumnErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {Principal}, 錯誤訊息: {principalColumnErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else
                {
                    WARNING($"NeedCheck, 輸入:{Principal}, 錯誤訊息: {principalColumnErrorWordings}");
                }
            }



            ///<summary>
            ///聯絡人欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 聯絡人");
            string[] ContactPersonArray = new string[] {"", "       ", "Ａｂ", "ab", "許功蓋ㄟㄟ", "許功蓋Bluce", "Ａ", "中", "英文ａｂｃ中文Ｄｅｆ符號#$%!@~&*^%", "中文字加英文字加數字_ＡＢＣＤＥＦＧＨＩＪ＿abcdefghij_1234567890_ABCDEFGHIJ_", "GoldPrinciple"};
            foreach (var contactPerson in ContactPersonArray)
            {
                IWebElement ContactPersonColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[2]/input"));
                ContactPersonColumn.Clear();
                ContactPersonColumn.SendKeys(contactPerson); // 填負責人
                string contactPersonColumnErrorWordings = driver.FindElement(By.Id("contactPerson-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnContactPerson= ContactPersonColumn.GetAttribute("value"); // 實際欄位讀到的值

                if(string.IsNullOrWhiteSpace(contactPerson) == true)
                {
                    if(contactPersonColumnErrorWordings == "必須填寫")
                    {
                        PASS($"輸入: {contactPerson}, 錯誤訊息: {contactPersonColumnErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {contactPerson}, 錯誤訊息: {contactPersonColumnErrorWordings}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else
                {
                    WARNING($"NeedCheck, 輸入:{contactPerson}, 錯誤訊息: {contactPersonColumnErrorWordings}");
                }
            }



            ///<summary>
            ///連絡電話-區碼-欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 連絡電話-區碼");
            string[] telephoneAreaCodeArray = new string[] {"", "       ", "12345", "ａｂｃａｂｃＤｅｆ", "abcdefgh", "中文字加英文數字", "!@#$%^&*(", "1232148123123", "55665566", "@英%^&文b字c", "１２３４", "12二", "91301434", "3a45", "6789", "02"};
            foreach (var areaCode in telephoneAreaCodeArray)
            {
                bool telephoneAreaCodeCheck = Regex.IsMatch(areaCode, @"^[0-9]*$");

                IWebElement TelephoneAreaCodeColumn  = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[1]"));
                TelephoneAreaCodeColumn.Clear();
                TelephoneAreaCodeColumn.SendKeys(areaCode); 
                string telephoneArea_error = driver.FindElement(By.Id("companyTelArea-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnAreaCodes = TelephoneAreaCodeColumn.GetAttribute("value"); // 實際欄位讀到的值

                if (string.IsNullOrWhiteSpace(areaCode) == true ) // 當輸入為空值
                {
                    if(telephoneArea_error == "必須填寫")
                    {
                        PASS($"輸入: {areaCode}, 錯誤訊息: {telephoneArea_error}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {areaCode}, 錯誤訊息: {telephoneArea_error}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else if (string.IsNullOrWhiteSpace(areaCode) != true && telephoneAreaCodeCheck != true) // 當輸入不是空值 & 輸入不為"全數字"
                {
                    if(telephoneArea_error == "僅能輸入數字")
                    {
                        PASS($"輸入: {areaCode}, 錯誤訊息: {telephoneArea_error}, 預期錯誤訊息: 僅能輸入數字");
                    }
                    else
                    {
                        FAIL($"輸入: {areaCode}, 錯誤訊息: {telephoneArea_error}, 預期錯誤訊息: 僅能輸入數字");
                    }
                }
                else if (telephoneAreaCodeCheck ==true) // 當輸入為全數字
                {
                    if (telephoneArea_error == "" && actualColumnAreaCodes.Length <=4) // 當沒有錯誤訊息 且 數字長度 < 4位
                    {
                         PASS($"輸入: {areaCode}, 輸入長度: {areaCode.Length}, 欄位實際讀到: {actualColumnAreaCodes}, 欄位長度: {actualColumnAreaCodes.Length}, 錯誤訊息: {telephoneArea_error}, 預期錯誤訊息: n/a, 預期長度: 4位數以內");
                    }
                    else
                    {
                        FAIL($"輸入: {areaCode}, 輸入長度: {areaCode.Length}, 欄位實際讀到: {actualColumnAreaCodes}, 欄位長度: {actualColumnAreaCodes.Length}, 錯誤訊息: {telephoneArea_error}, 預期錯誤訊息: n/a, 預期長度: 4位數以內");
                    }
                }
                else // 以上Case皆不是時, 需人工check
                {
                       WARNING($"NeedCheck, 輸入: {areaCode}, 欄位實際讀到: {actualColumnAreaCodes}, 欄位長度: {actualColumnAreaCodes.Length}, 錯誤訊息: {telephoneArea_error}");
                }
            }
            ///<summary>
            ///連絡電話-主碼-欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 連絡電話-主碼");
            string[] telephoneMainArray = new string[] {"", "       ", "12345", "ａｂｃａｂｃＤｅｆ", "ABCDEFG", "abcdefgh", "中文字加英文數字", "!@#$%^&*(", "１２３４５６７８", "123４５６78", "55665566", "@英%^&文b字c", "１２３４", "12二", "67890123", "3a4567b8", "6789", "2345678923456789", "8825252"};
            foreach (var telephoneMain in telephoneMainArray )
            {
                bool telephoneMainCheck = Regex.IsMatch(telephoneMain, @"^[0-9]*$");

                IWebElement TelephoneMainColumn  = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[2]"));
                TelephoneMainColumn.Clear();
                TelephoneMainColumn.SendKeys(telephoneMain); 
                string telephoneMain_error = driver.FindElement(By.Id("tel-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnTelephoneMain = TelephoneMainColumn.GetAttribute("value"); // 實際欄位讀到的值

                if (string.IsNullOrWhiteSpace(telephoneMain) == true ) // 當輸入為空值
                {
                    if(telephoneMain_error == "必須填寫")
                    {
                        PASS($"輸入: {telephoneMain}, 錯誤訊息: {telephoneMain_error}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneMain}, 錯誤訊息: {telephoneMain_error}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else if (string.IsNullOrWhiteSpace(telephoneMain) != true && telephoneMainCheck != true) // 當輸入不為空值但輸入"不全為數字"
                {
                    if(telephoneMain_error == "僅能輸入數字")
                    {
                        PASS($"輸入: {telephoneMain}, 錯誤訊息: {telephoneMain_error}, 預期錯誤訊息: 僅能輸入數字");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneMain}, 錯誤訊息: {telephoneMain_error}, 預期錯誤訊息: 僅能輸入數字");
                    }
                }
                else if (telephoneMainCheck ==true) // 當輸入為全數字
                {
                    if (telephoneMain_error == "" && actualColumnTelephoneMain.Length <=8) // 當沒有錯誤訊息 且 數字長度 <= 8位
                    {
                         PASS($"輸入: {telephoneMain}, 輸入長度: {telephoneMain.Length}, 欄位實際讀到: {actualColumnTelephoneMain}, 欄位長度: {actualColumnTelephoneMain.Length}, 錯誤訊息: {telephoneMain_error}, 預期錯誤訊息: n/a, 預期長度: 8 位數以內");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneMain}, 輸入長度: {telephoneMain.Length}, 欄位實際讀到: {actualColumnTelephoneMain}, 欄位長度: {actualColumnTelephoneMain.Length}, 錯誤訊息: {telephoneMain_error}, 預期錯誤訊息: n/a, 預期長度: 8 位數以內");
                    }
                }
                else // 以上Case皆不是時, 需人工check
                {
                       WARNING($"NeedCheck, 輸入: {telephoneMain}, 欄位實際讀到: {actualColumnTelephoneMain}, 欄位長度: {actualColumnTelephoneMain.Length}, 錯誤訊息: {telephoneMain_error}");
                }
            }
            ///<summary>
            ///連絡電話-分機-欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核: 連絡電話-分機");
            string[] telephoneExtensionArray = new string[] {"", "       ", "12345", "ａｂｃａｂｃＤｅｆ", "ABCDEFG", "abcdefgh", "中文字加英文數字", "!@#$%^&*(", "１２３４５６７８", "123４５６78", "55665566", "@英%^&文b字c", "１２３４", "12二", "67890123", "3a4567b8", "6789", "2345678923456789", "8825252", "123", "131517131517", "9527"};
            foreach (var telephoneExtension in telephoneExtensionArray )
            {
                bool telephoneExtensionCheck = Regex.IsMatch(telephoneExtension, @"^[0-9]*$");

                TelephoneExtensionColumn.Clear();
                TelephoneExtensionColumn.SendKeys(telephoneExtension); 
                string telephoneExt_error = driver.FindElement(By.Id("telext-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnTelephoneExtension = TelephoneExtensionColumn.GetAttribute("value"); // 實際欄位讀到的值

                if (string.IsNullOrWhiteSpace(telephoneExtension) == true ) // 當輸入為空值
                {
                    if(telephoneExt_error == "必須填寫")
                    {
                        PASS($"輸入: {telephoneExtension}, 錯誤訊息: { telephoneExt_error }, 預期錯誤訊息: n/a (非必填)");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneExtension}, 錯誤訊息: { telephoneExt_error }, 預期錯誤訊息: n/a (非必填)");
                    }
                }
                else if (string.IsNullOrWhiteSpace(telephoneExtension) != true && telephoneExtensionCheck  != true) // 當輸入不為空值但輸入"不全為數字"
                {
                    if(telephoneExt_error  == "僅能輸入數字")
                    {
                        PASS($"輸入: {telephoneExtension}, 錯誤訊息: {telephoneExt_error }, 預期錯誤訊息: 僅能輸入數字");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneExtension}, 錯誤訊息: {telephoneExt_error }, 預期錯誤訊息: 僅能輸入數字");
                    }
                }
                else if ( telephoneExtensionCheck == true  && telephoneExtension.Length < 6 )  // 當輸入為全數字 且 數字長度 < 6位
                {
                    if (telephoneExt_error == "") // 當沒有錯誤訊息時
                    {
                        PASS($"輸入: {telephoneExtension}, 輸入長度: {telephoneExtension.Length}, 欄位實際讀到: {actualColumnTelephoneExtension}, 欄位長度: {actualColumnTelephoneExtension.Length}, 錯誤訊息: {telephoneExt_error }, 預期錯誤訊息: n/a, 預期長度: 6 位數以內");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneExtension}, 輸入長度: {telephoneExtension.Length}, 欄位實際讀到: {actualColumnTelephoneExtension}, 欄位長度: {actualColumnTelephoneExtension.Length}, 錯誤訊息: {telephoneExt_error }, 預期錯誤訊息: n/a, 預期長度: 6 位數以內");
                    }
                }
                else if (telephoneExtensionCheck ==true  &&  telephoneExtension.Length >= 6) // 當輸入為全數字 且 數字長度 >= 6位
                {
                     if (actualColumnTelephoneExtension.Length == 6) // 當欄位實際長度為6位
                    {
                        PASS($"輸入: {telephoneExtension}, 輸入長度: {telephoneExtension.Length}, 欄位實際讀到: {actualColumnTelephoneExtension}, 欄位長度: {actualColumnTelephoneExtension.Length}, 錯誤訊息: {telephoneExt_error }, 預期錯誤訊息: n/a, 預期長度: 6 位數以內");
                    }
                    else
                    {
                        FAIL($"輸入: {telephoneExtension}, 輸入長度: {telephoneExtension.Length}, 欄位實際讀到: {actualColumnTelephoneExtension}, 欄位長度: {actualColumnTelephoneExtension.Length}, 錯誤訊息: {telephoneExt_error }, 預期錯誤訊息: n/a, 預期長度: 6 位數以內");
                    }
                }
            }



            ///<summary>
            /// 行動電話號碼欄位檢核
            ///</summary>
            INFO("");
            INFO("欄位檢核:行動電話號碼");
            string[] phones = new string[] { "", "          ","!@#$%^", "許功蓋", "ABCDEF", "0897654321", "A987654321", "０987654321", "098765432１", "0987 543 1", "0901234567890", "0910313414" };
            foreach (var input in phones)
            {
                // int string_length = System.Text.Encoding.Default.GetBytes(input).Length; //  擷取字串長度bytes (UTF-8標準, 半形英數字 = 1 byte, 中文&全形英數字 = 3 bytes)
                bool numericCheck = Regex.IsMatch(input, @"^[0-9]*$"); // 判斷輸入是否為全數字
                bool cellPhoneNumberCheck = Regex.IsMatch(input, @"^09\d{8}$"); // 判斷輸入是否為行動電話09開頭格式 (正則表示式: 09開頭後面8碼數字) 
                
                IWebElement CellphoneColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[6]/td[2]/input"));
                CellphoneColumn.Clear();
                CellphoneColumn.SendKeys(input);
                string cellphone_error = driver.FindElement(By.Id("cellPhone-error")).Text; // 欄位檢核錯誤訊息
                string actualColumnCellphone = CellphoneColumn.GetAttribute("value"); // 欄位實際讀到的值
                 
                 if (string.IsNullOrWhiteSpace(input) == true ) // 當輸入為空值
                {
                    if(cellphone_error== "必須填寫")
                    {
                        PASS($"輸入: {input}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: 必須填寫");
                    }
                    else
                    {
                        FAIL($"輸入: {input}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: 必須填寫");
                    }
                }
                else if (numericCheck != true) // 當輸入"不全為數字"
                {
                    if(cellphone_error == "行動電話格式錯誤")
                    {
                        PASS($"輸入: {input}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: 行動電話格式錯誤");
                    }
                    else
                    {
                        FAIL($"輸入: {input}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: 行動電話格式錯誤");
                    }
                }
                else if (numericCheck == true && cellPhoneNumberCheck != true) // 當輸入為全數字 但 非行動電話格式
                {
                    if (cellphone_error == "行動電話格式錯誤") 
                    {
                        PASS($"輸入: {input}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: 行動電話格式錯誤");
                    }
                    else
                    {
                        FAIL($"輸入: {input}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: 行動電話格式錯誤");
                    }
                }
                 else if (cellPhoneNumberCheck == true)
                {
                    PASS($"輸入: {input}, 輸入長度: {input.Length}, 欄位實際讀到: {actualColumnCellphone}, 欄位長度: {actualColumnCellphone.Length}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: n/a, 預期長度: 09開頭的10位數字");
                }
                else // 以上Case皆不是時, 需人工check
                {
                       WARNING($"NeedCheck, 輸入: {input}, 輸入長度: {input.Length}, 欄位實際讀到: {actualColumnCellphone}, 欄位長度: {actualColumnCellphone.Length}, 錯誤訊息: {cellphone_error}, 預期錯誤訊息: n/a, 預期長度: 09開頭的10位數字");
                }
            }



            ///<summary>
            /// 洽詢紓困專案下拉選單
            ///</summary>
            INFO("");
            INFO("檢查:紓困專案下拉選單選項");
            int rescueProjectCounts = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li")).Count; // 選項數量
            for (int i = 2; i < rescueProjectCounts; i++)
			{
                IWebElement RescueProjectDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li")); // 紓困專案下拉選單
                RescueProjectDropDownList.Click();
                IWebElement SelectResuceProject = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li[{i}]"));
                SelectResuceProject.Click();
                WARNING($"紓困方案 {i - 1}: {RescueProjectDropDownList.Text}");
			}



            ///<summary>
            /// 縣市&分行欄位
            ///</summary>
            INFO("");
            INFO("檢查:縣市&分行下拉選單選項");
            IWebElement Country_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li"));
            Country_DropDownList.Click(); //展開"請選擇縣市"下拉選單
            IWebElement SelectCountry = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li/ul/li[2]/span"));
            SelectCountry.Click(); // 點選"台北市"
            IWebElement Branch_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]"));
            Branch_DropDownList.Click(); // 展開"分行"下拉選單
            IWebElement SelectBranch = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li[2]/span"));
            SelectBranch.Click(); //點選"營業處"



            ///<summary>
            /// 蒞行日期&時間欄位
            ///</summary>
            INFO("");
            INFO("檢查:蒞行日期&時間下拉選單選項");
            Tools.ScrollPageUpOrDown(driver, 700);
             int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li")).Count; // 獲取蒞行日期下拉選單選項數量
             int time_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li")).Count; // 獲取蒞行時間下拉選單選項數量

            for (int m= 2; m < date_amount; m++)
			{
                IWebElement DateDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li")); 
                DateDropDownList.Click();  // 展開蒞行日期下拉選單
                IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li[{m}]/span")); 
                SelectDate.Click(); // 點選一個"日期"
                WARNING($"第 {m - 1} 個可選日期: {DateDropDownList.Text}");
			}
             for (int n= 2; n < time_amount; n++)
			{
                IWebElement TimeDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li"));
                TimeDropDownList .Click(); // 展開"時段"下拉選單
                IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li[{n}]/span"));
                SelectTime.Click(); // 點選一個"時段"
                WARNING($"第 {n - 1} 個可選日期: {TimeDropDownList.Text}");
			}



            ///<summary>
            /// 檢查 "預約蒞行時間" 說明文字
            ///</summary>
           INFO("");
           INFO("檢查:預約蒞行時間說明文字");
           string appointment_visit_time_column = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]")).Text;
           string appointment_visit_time_information = appointment_visit_time_column.Substring(appointment_visit_time_column.IndexOf("※"));
            if (appointment_visit_time_information == "※為提升服務品質並進行顧客蒞行服務分流，本服務僅供每人每次預約1個時段，造成不便，請您見諒。")
            {
                PASS("符合預期");
            }
            else
            {
                FAIL($"錯誤訊息: {appointment_visit_time_information}, 預期錯誤訊息: ※為提升服務品質並進行顧客蒞行服務分流，本服務僅供每人每次預約1個時段，造成不便，請您見諒。");
            }


            ///<summary>
            /// 檢查 "我已閱讀並同意" 錯誤訊息
            ///</summary>
           INFO("");
           INFO("檢查:我已閱讀並同意錯誤訊息");
           IWebElement IHaveReadRadioButton = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[11]/td[2]/div/a"));
           
        rerun_ihaveread:
            string i_have_read_status = IHaveReadRadioButton.GetAttribute("class"); // 我已閱讀 radio button 狀態 (checked: 已勾選)
            string i_have_read_error = driver.FindElement(By.Id("apply-error")).Text; // 我已閱讀欄位錯誤訊息
            if (i_have_read_status != "checked")
            { 
                if (i_have_read_error == "必須填寫")
                {
                    PASS($"未勾選錯誤訊息: {i_have_read_error}, 預期錯誤訊息: 必須填寫");
                }
                else
                {
                    FAIL($"未勾選錯誤訊息: {i_have_read_error}, 預期錯誤訊息: 必須填寫");
                }
                IHaveReadRadioButton.Click(); // 點"我已閱讀" radio button
                 goto rerun_ihaveread;
            }
            else
            {
                if (i_have_read_error == "")
                {
                    PASS($"已勾選錯誤訊息: {i_have_read_error}, 預期錯誤訊息: n/a");
                }
                else
                {
                    FAIL($"已勾選錯誤訊息: {i_have_read_error}, 預期錯誤訊息: n/a");
                }
            }


            ///<summary>
            /// 檢查 "圖型驗證碼" 錯誤訊息
            /// </summary>
            INFO("");
            INFO("檢查: 圖型驗證碼錯誤訊息");
            Tools.ScrollPageUpOrDown(driver, 800);
            IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath("//*[@id='captchaValue']")); // 驗證碼輸入欄位
        rerun_CaptchaSteps:
            string image_verification_column_value = driver.FindElement(By.Id("captchaValue")).GetAttribute("value"); // 驗證碼欄位實際讀到的值
            string image_verification_error = driver.FindElement(By.Id("captchaWrong")).Text; // 驗證碼錯誤訊息
            if (string.IsNullOrWhiteSpace(image_verification_column_value) == true)
            {
                if (image_verification_error == "請輸入驗證碼")
                {
                     PASS($"未輸入, 錯誤訊息: {image_verification_error}, 預期錯誤訊息: 請輸入驗證碼");
                }
                else
                {
                    FAIL($"未輸入, 錯誤訊息: {image_verification_error}, 預期錯誤訊息: 請輸入驗證碼");
                }    
                ImageVerificationCodeColumn.SendKeys("9527");
                SubmitButton.Click();
                goto rerun_CaptchaSteps;
            }
            else
            {
                if (image_verification_error == "驗證碼錯誤")
                {
                     PASS($"輸入錯誤的驗證碼, 錯誤訊息: {image_verification_error}, 預期錯誤訊息: 驗證碼錯誤");
                }
                else
                {
                    FAIL($"FAIL, 輸入錯誤的驗證碼, 錯誤訊息: {image_verification_error}, 預期錯誤訊息: 驗證碼錯誤");
                }  
            }



            ///<summary>
            ///檢查個資法告知事項超連結
            ///</summary>
            INFO("");
            INFO("檢查: 個資法告知事項超連結");
            string personal_indormation_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[11]/td[2]/a")).GetAttribute("href"); // 獲取個資法告知事項的超連結
            if (personal_indormation_hyperlink == "https://www.esunbank.com.tw/bank/about/announcement/privacy/privacy-statement")
            {
                PASS("個資法告知事項超連結 符合預期");
            }
            else
            {
                FAIL($"個資法告知事項 link: {personal_indormation_hyperlink}, 預期 link: https://www.esunbank.com.tw/bank/about/announcement/privacy/privacy-statement");
            }


            ///<summary>
            ///檢查下方文案區提示文案
            ///</summary>
            INFO("");
            INFO("檢查: 下方文案區提示文案");
            string promot_wordings = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[3]")).Text;
            if (promot_wordings == "※提醒您，點擊「確認」送出預約後，請依預約前往分行臨櫃申請，若您未於預約時間來行，為配合政府防疫措施，請勿直接至分行申請。")
            {
                PASS("下方文案區提示文案 符合預期");
            }
            else
            {
                FAIL($"下方文案區提示文案:  實際敘述: {promot_wordings}, 預期敘述: ※提醒您，點擊「確認」送出預約後，請依預約前往分行臨櫃申請，若您未於預約時間來行，為配合政府防疫措施，請勿直接至分行申請。");
            }

            CloseBrowser();
        }
    }
}


