using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using OpenQA.Selenium.Support.UI;
using System.IO;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.BailoutLoan
{
    public class ���~�V�x�U�ڹw���A��_�e�X���:IntegrationTestBase
    {
        public ���~�V�x�U�ڹw���A��_�e�X���(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/small-business/tools/apply/sbloan-appointment";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void �e�X���(string browser)
        {
            StartTestCase(browser, "���~�V�x�U�ڹw���A��_�Ҧ������ưe�X", "York");

            Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha");
            System.Threading.Thread.Sleep(100);
            Tools.CleanUPFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha"); //�M��captcha��Ƨ�

            int country_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li/ul/li")).Count; // ���������total������
            for (int country_index = 1; country_index <= country_amount - 1; country_index++)
            {
                IWebElement Country_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li"));
                Country_DropDownList.Click(); // �i�} "�п�ܿ���" �U�Կ��

                string country_xpath = $"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li[{country_index + 1}]";
                IWebElement SelectCountry = driver.FindElement(By.XPath(country_xpath));
                SelectCountry.Click(); // �I��@��"����"

                int branch_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li")).Count;
                for (int branch_index = 1; branch_index <= branch_amount - 1; branch_index++) // index 1 = �Ĥ@�Ӥ���
                {
                    if (branch_index >= 2)
                    {
                        driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li")).Click(); // ���s�I�������
                        driver.FindElement(By.XPath(country_xpath)).Click(); // ���s�￤��
                    }
                    IWebElement Branch_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li[2]/span")); // �i�} "����" �U�Կ��
                    Branch_DropDownList.Click(); // �i�}"����"�U�Կ��
                    string branch_xpath = $"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li[{branch_index + 1}]/span";
                    IWebElement SelectBranch = driver.FindElement(By.XPath(branch_xpath));
                    SelectBranch.Click(); //�I��@��"����



                    /// <summary>
                    /// �񤽥q�W��
                    ///  </summary>
                    IWebElement CompanyNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]/input"));
                    CompanyNameColumn.Clear();
                    CompanyNameColumn.SendKeys($"AutomatedCompany {country_index}-{branch_index}"); // �񤽥q�W��



                    /// <summary>
                    /// �񤽥q�νs
                    ///  </summary>
                    IWebElement CompanyUniformNumbersColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]/input"));
                    CompanyNameColumn.Clear();
                    CompanyNameColumn.SendKeys(Tools.CreateUniformNumber()); // �񤽥q�νs



                    /// <summary>
                    /// ��t�d�H �p���H
                    ///  </summary>
                    IWebElement PrincipleColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[2]/input"));
                    IWebElement ContactPersonColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[2]/input"));
                    PrincipleColumn.Clear();
                    PrincipleColumn.SendKeys($"Boss");
                    ContactPersonColumn.Clear();
                    ContactPersonColumn.SendKeys($"LittleBrother");



                    /// <summary>
                    /// ��s���q��: �ϽX+�D�X+����
                    ///  </summary>
                    IWebElement TelephoneAreaCodeColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[1]"));
                    IWebElement TelephoneMainColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[2]"));
                    IWebElement TelephoneExtensionColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[3]"));
                    TelephoneAreaCodeColumn.Clear();
                    TelephoneAreaCodeColumn.SendKeys($"02");
                    TelephoneMainColumn.Clear();
                    TelephoneMainColumn.SendKeys($"8825252");
                    TelephoneExtensionColumn.Clear();
                    TelephoneExtensionColumn.SendKeys($"9527");



                    /// <summary>
                    /// ���ʹq��
                    ///  </summary>
                    IWebElement CellPhoneColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[1]"));
                    CellPhoneColumn.Clear();
                    CellPhoneColumn.SendKeys(Tools.CreateCellPhoneNumber());



                    ///<summary>
                    /// �����V�x�M�פU�Կ��
                    ///</summary>
                    IWebElement RescueProjectDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li")); // �V�x�M�פU�Կ��
                    RescueProjectDropDownList.Click();
                    int rescueProjectCounts = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li")).Count; // �ﶵ�ƶq
                    Random rescueOption = new Random();
                    int projectSelected = rescueOption.Next(2, rescueProjectCounts + 1);
                    string selectRescueProjectXPath = $"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li[{projectSelected}]";
                    IWebElement SelectResuceProject = driver.FindElement(By.XPath(selectRescueProjectXPath));
                    SelectResuceProject.Click(); // �H������V�x�M��



                    ///<summary>
                    /// �H����ܻY����&�ɶ�
                    ///</summary>
                    IWebElement DateDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li"));
                    IWebElement TimeDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li"));
                    Tools.ScrollPageUpOrDown(driver, 700);
                    int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li")).Count; // ����Y�����U�Կ��ﶵ�ƶq
                    int time_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li")).Count; // ����Y��ɶ��U�Կ��ﶵ�ƶq

                    Random ran_date = new Random();
                    Random ran_time = new Random();
                    int ran_date_index = ran_date.Next(2, date_amount);
                    int ran_time_index = ran_time.Next(2, time_amount);

                    DateDropDownList.Click();  // �i�}"�Y����"�U�Կ��
                    IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li[{ran_date_index}]/span"));
                    SelectDate.Click(); // �H���I��@��"���"
                    TimeDropDownList.Click(); // �i�}"�ɬq"�U�Կ��
                    IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li[{ran_time_index}]/span"));
                    SelectTime.Click(); // �H���I��@��"�ɬq"



                    ///<summary>
                    /// �I "�ڤw�\Ū�æP�N"
                    ///</summary>
                    IWebElement IHaveReadRadioButton = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[11]/td[2]/div/a"));
                    IHaveReadRadioButton.Click();



                    ///<summary>
                    /// �I"�T�{"
                    ///</summary>
                    IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));
                    SubmitButton.Click(); // �I �e�X



                    ///<summary>
                    /// �۰ʿ�J�Ϥ����ҽX
                    ///</summary>
                    int verify_count = 1; // verify_count: ����retry���ҽX����
                retryagain:
                    IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='ImgCaptcha']")); // Captcha�Ϥ���m
                    IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath("//*[@id='captchaValue']")); // ��J���ҽX���

                    Tools.ElementSnapshotshot(CaptchaPicture, $@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //snapshot���ҽX�Ϥ�
                    string verify_code_result = TestBase.TesseractOCRIdentify($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png", 0.75); //�ѪR�X���ҽX

                    if (verify_count >= 10) // �̧ǧR���ª�picture
                    {
                        File.Delete($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count - 9}.png");
                    }
                    ImageVerificationCodeColumn.Clear();
                    ImageVerificationCodeColumn.SendKeys(verify_code_result); // ��J���ҽX

                    if (driver.FindElement(By.Id("captchaWrong")).Text == "") // �ˬd���ҽX���~�T�����O�_����
                    {
                        SubmitButton.Click(); // �I"�T�{"button

                        if (driver.FindElement(By.Id("captchaWrong")).Text == "")
                        {
                            goto gotoneststep;
                        }
                        else
                        {
                            CaptchaPicture.Click();
                            verify_count++;
                            goto retryagain;
                        }
                    }
                    else
                    {
                        CaptchaPicture.Click();
                        verify_count++;
                        goto retryagain;
                    }

                gotoneststep:
                    WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                    wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // ���ݪ���ݨ�q������ 
                    System.Threading.Thread.Sleep(1000);
                    Tools.ScrollPageUpOrDown(driver, 700);
                    string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\BailoutLoan";
                    Tools.CreateSnapshotFolder(snapshotpath);
                    Tools.FullScreenshot($@"{snapshotpath}\��{country_index}������{branch_index}����_�V�x�M��{projectSelected - 1}_�w����{ran_date_index - 1}���{ran_time_index - 1}�ɬq.png"); //��@�I��

                    driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // �I�q������ "X" ���s
                    System.Threading.Thread.Sleep(1000);

                    driver.Navigate().GoToUrl(testurl);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);

                }
            }
            driver.Quit();
        }
    }
}


