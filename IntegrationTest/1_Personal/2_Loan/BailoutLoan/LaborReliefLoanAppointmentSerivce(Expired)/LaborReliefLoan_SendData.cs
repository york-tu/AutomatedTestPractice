using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using OpenQA.Selenium.Support.UI;
using System.IO;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Personal.Loan.BailoutLoan
{
    public class �Ҥu�V�x�U�ڹw���A��_�e�X���: IntegrationTestBase
    {
        public �Ҥu�V�x�U�ڹw���A��_�e�X���(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/labor-loan-appointment#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void �e�X���(string browser)
        {
            StartTestCase(browser, "�Ҥu�V�x�U�ڹw���A��_�Ҧ������ưe�X", "York");

            Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha");
            System.Threading.Thread.Sleep(100);
            Tools.CleanUPFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha"); //�M��captcha��Ƨ�

            for (int i = 2; i <= 21; i++) // initial i =2
                {

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 35, 1, 31, 1, 3, 3, 11, 3, 14, 2, 1, 2, 1, 1, 10, 14, 1, 2, 1, 1 }; //�������������
                    while (j <= arrray[i - 2])
                    {
                        string timesavepath = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");

                        // �w�q���XPath
                        string Cuntry_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{i}]/span";
                        string Branch_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{j + 1}]/span";


                        IWebElement FullNameColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.name_column_Xpath()));
                        FullNameColumn.Clear();
                        FullNameColumn.SendKeys($"�����H{i - 1}_{j}"); // ��m�W


                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.ID_column_XPath()));
                        IdentityCardColumn.Clear();
                        bool sex = false;
                        if (j % 2 == 0)
                            sex = true;
                        IdentityCardColumn.SendKeys(Tools.CreateIDNumber(sex, j % 21)); // �ޥ�tools�������Ҧr�����;� & �񨭤������


                        IWebElement CellPhoneColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.cellphone_column_XPath()));
                        CellPhoneColumn.Clear();
                        CellPhoneColumn.SendKeys(Tools.CreateCellPhoneNumber()); // �ޥ�tools���q�ܸ��X���;� & ��q�ܸ��X���


                        IWebElement BirthdayCalendarIcon = driver.FindElement(By.Id("datepicker1-button"));
                        BirthdayCalendarIcon.Click(); // �I�}�X�ͦ~���p��� 
                        Random Year = new Random();
                        Random Month = new Random();
                        // IWebElement Year_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]"));
                        IWebElement SelectYear = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[1]/select/option[{Year.Next(1, 121)}]")); // �H���~��
                        SelectYear.Click();
                        //  IWebElement Month_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]"));
                        IWebElement SelectMonth = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[2]/select/option[{Month.Next(1, 12)}]")); // �H�����
                        SelectMonth.Click(); // ����

                    rerun:
                        Random week = new Random();
                        Random day = new Random();
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr[{week.Next(1, 5)}]/td[{day.Next(1, 7)}]")); // �H�����
                        string date_value = date.GetAttribute("class");
                        if (date_value == "is-empty") // �P�_����쬰�Ů�(��ܷ��S���o��l), ���s�����H�����
                        {
                            goto rerun;
                        }
                        else
                        {
                            date.Click();
                        }

                        IWebElement Country_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.country_dropdownlist_XPath()));
                        Country_DropDownList.Click(); //�i�}"�п�ܿ���"�U�Կ��

                        IWebElement SelectCountry = driver.FindElement(By.XPath(Cuntry_Xpath));
                        SelectCountry.Click(); // �I��@��"����"

                        IWebElement Branch_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.branch_dropdownlist_XPath()));
                        Branch_DropDownList.Click(); // �i�}"����"�U�Կ��

                        IWebElement SelectBranch = driver.FindElement(By.XPath(Branch_Xpath));
                        SelectBranch.Click(); //�I��@��"����"


                        IWebElement Date_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.date_dropdownlist_XPath()));
                        int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li")).Count;
                        int time_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li")).Count;
                        Random ran_date = new Random();
                        Random ran_time = new Random();

                        Date_DropDownList.Click(); //�i�}"�Y����"�U�Կ��
                        IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li[{ran_date.Next(2, date_amount)}]/span"));
                        SelectDate.Click(); // �`���I��@��"���"
                        IWebElement Time_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.time_dropdownlist_XPath()));
                        Time_DropDownList.Click(); // �i�}"�ɬq"�U�Կ��
                        IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li[{ran_time.Next(2, time_amount)}]/span"));
                        SelectTime.Click(); // �`���I��@��"�ɬq"

                        IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath(LaborReliefLoan_XPath.i_have_read_button_XPath()));
                        IHaveReadRadioButtin.Click(); // �I"�ڤw�\Ū" radio button

                        IWebElement SubmitButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.submit_button_XPath()));

                        IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='ImgCaptcha']")); //�Ϥ����

                       

                        int verify_count = 1; // verify_count: ����retry���ҽX����
                    retryagain:
                        IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.image_verify_code_column_XPath())); // ��J���ҽX���

                        Tools.SCrollToElement(driver, FullNameColumn);
                        Tools.ElementSnapshotshot(CaptchaPicture, $@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //snapshot���ҽX�Ϥ�

                        if (verify_count >=10) // �̧ǧR���ª�picture
                        {
                            File.Delete($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count-9}.png");
                        }

                        string verify_code_result = TestBase.TesseractOCRIdentify($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png", 0.75); //�ѪR�X���ҽX

                        // Tools.ElementTakeScreenShot(CaptchaPicture, $@"{snapshotpath}\ImageVerifyCode.png");
                        // string verify_code_result = Tools.IronOCR($@"{snapshotpath}\ImageVerifyCode.png"); //�ѪR�X���ҽX: Iron_OCR
                        // string verify_code_result = Tools.BaiduOCR($@"{snapshotpath}\ImageVerifyCode.png"); //�ѪR�X���ҽX: Baidu_OCR

                        ImageVerificationCodeColumn.Clear();
                        ImageVerificationCodeColumn.SendKeys(verify_code_result); // ��J���ҽX
                        if (driver.FindElement(By.Id("captchaWrong")).Text =="") // �ˬd���ҽX���~�T�����O�_����
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

                        WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                        wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // ���ݪ���ݨ�q������ 
                        System.Threading.Thread.Sleep(300);

                        Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\LaborReliefLoan");
                        Tools.PageSnapshot($@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\LaborReliefLoan\��{i - 1}������{j}����_�ӽ�_��{ran_date}���{ran_time}�ɬq.png", driver); //��@�I��


                        driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // �I�q������ "X" ���s
                        System.Threading.Thread.Sleep(3000);

                        driver.Navigate().GoToUrl(testurl);
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);

                        j++;
                    }
                }
            CloseBrowser();
            }
        }
    }


