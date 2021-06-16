using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Xunit;
using References;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;


namespace LaborReliefLoanAppointmentServiceTest
{
    public class �Ҥu�V�x�U�ڹw���A��_ErrorHandling_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/labor-loan-appointment#";


        [Theory]
        [InlineData(BrowserType.Chrome)]
        [InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                driver.Manage().Window.Maximize();


                for (int i = 2; i <= 21; i++) // initial i =2
                {

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 35,1, 31, 1, 3, 3, 11, 3, 14, 2, 1, 2, 1, 1, 10, 14, 1, 2, 1, 1}; //�������������
                    while (j <= arrray[i - 2])
                    {
                        string timesavepath = System.DateTime.Now.ToString("yyyy-MM-dd_hhmm");

                        // �w�q���XPath
                        string Cuntry_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{i}]/span";
                        string Branch_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{j+1}]/span";

                        IWebElement SubmitButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.submit_button_XPath()));
                        SubmitButton.Click(); // �I"�T�{"button

                        IWebElement FullNameColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.name_column_Xpath()));
                        FullNameColumn.Clear();
                        FullNameColumn.SendKeys($"�����H{i-1}_{j}"); // ��m�W


                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.ID_column_XPath()));
                        IdentityCardColumn.Clear();

                        bool sex = false;
                        if (j % 2 == 0)
                            sex = true;
                        IdentityCardColumn.SendKeys(Tools.CreateIDNumber(sex, j%21)); // �ޥ�tools�������Ҧr�����;� & �񨭤������


                        IWebElement CellPhoneColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.cellphone_column_XPath()));
                        string[] phones = new string[] { "ABCDEF", "0897654321", "", $"{Tools.CreateCellPhoneNumber()}" };
                        foreach (var input in phones)
                        {
                            CellPhoneColumn.Clear();
                            bool CellPhoneNumberCheck = Regex.IsMatch(input, @"^09\d{ 8}$");
                            string cellphone_error = driver.FindElement(By.Id("cellPhone-error")).GetAttribute("value");
                            CellPhoneColumn.SendKeys(input);

                            if (input == "")
                            {
                                Assert.Equal("������g", input);
                            }
                            else if (CellPhoneNumberCheck != true)
                            {
                                Assert.Equal("��ʹq�ܮ榡���~", input);
                            }
                            else
                            {
                                Assert.Equal("", input);
                            }
                        }


                        IWebElement BirthdayCalendarIcon = driver.FindElement(By.Id("datepicker1-button"));
                        BirthdayCalendarIcon.Click(); // �I�}�X�ͦ~���p��� 
                        Random Year = new Random();
                        Random Month = new Random();
                        // IWebElement Year_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]"));
                        IWebElement SelectYear = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[1]/select/option[{Year.Next(1,121)}]")); // �H���~��
                        SelectYear.Click();
                       //  IWebElement Month_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]"));
                        IWebElement SelectMonth = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[2]/select/option[{Month.Next(1, 12)}]")); // �H�����
                        SelectMonth.Click(); // ����

                        rerun:
                        Random week = new Random();
                        Random day = new Random();
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr[{week.Next(1,5)}]/td[{day.Next(1,7)}]")); // �H�����
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


                        for (int m = 2; m <=12; m++)
                        {
                            for (int n = 2; n <=6; n++)
                            {
                                IWebElement Date_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.date_dropdownlist_XPath()));
                                Date_DropDownList.Click(); //�i�}"�Y����"�U�Կ��
                                IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li[{m}]/span"));
                                SelectDate.Click(); // �I��@��"���"
                                IWebElement Time_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.time_dropdownlist_XPath()));
                                Time_DropDownList.Click(); // �i�}"�ɬq"�U�Կ��
                                IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li[{n}]/span"));
                                SelectTime.Click(); // �I��@��"�ɬq"
                                
                                string ErrorMessage = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/div[1]")).Text;
                                if (ErrorMessage == "")
                                {
                                    string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\LaborReliefLoan";
                                    Tools.CreateSnapshotFolder(snapshotpath);
                                    System.Threading.Thread.Sleep(100);
                                    Tools.TakeScreenShot($@"{snapshotpath}\��{i - 1}������{j}����_�ӽ�_��{m}���{n}�ɬq.png", driver); //��@�I��
                                }

                            }
                        }


                        //Tools.SCrollToElement(driver, CellPhoneColumn);
                        //IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.image_verify_code_column_XPath()));
                        //ImageVerificationCodeColumn.Click(); // �I"��J���ҽX"���
                        //ImageVerificationCodeColumn.Clear();
                        //System.Threading.Thread.Sleep(15000); // ���ݿ�J ���ҽX(����15��)

                        //IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath(LaborReliefLoan_XPath.i_have_read_button_XPath()));
                        //IHaveReadRadioButtin.Click(); // �I"�ڤw�\Ū" radio button

                        

                         
                        j++;
                     }
                }
                driver.Quit();
            }
        }
    }
}
    

