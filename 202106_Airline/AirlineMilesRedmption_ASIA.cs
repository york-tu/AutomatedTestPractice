using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Xunit;
using References;
using System;
using System.IO;
using System.Globalization;
using System.Linq;
using IronXL;
using System.Drawing;
using System.Windows.Forms;
using CsvHelper;
using CSVHeader;


namespace AirlineMilesRedemptionTest
{
    public class AirlineMilesRedemption // ��Ũ��{�I������
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/credit-card/reward/transit/asiamiles?dev=mobile";
        private readonly string Version = "Mobile";
        private readonly string excel_savepath = @"D:\����I�ƧI��TestReport.xlsx";
        private readonly string testcase_name = "�Ȭw�U���q";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void AMR_AISA_M(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");

                WorkBook xlsWorkbook;
                WorkSheet xlsSheet;

                if (File.Exists(excel_savepath) == true && WorkBook.Load(excel_savepath).GetWorkSheet(testcase_name) == null)
                { // �P�_��excel�ɦs�b �� �S��"xxx"�u�@�� >>> Ū����excel��, new create "xxx" �u�@�� 
                    xlsWorkbook = WorkBook.Load(excel_savepath); 
                    xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name); 
                }
                else if (File.Exists(excel_savepath) == true && WorkBook.Load(excel_savepath).GetWorkSheet(testcase_name) != null)
                { // �P�_��excel�ɦs�b �B ��"xxx"�u�@�� >>> Ū����excel��, Ū��"xxx" �u�@��
                    xlsWorkbook = WorkBook.Load(excel_savepath);
                    xlsSheet = xlsWorkbook.GetWorkSheet(testcase_name); 
                }
                else
                { // �P�_��excel�ɤ��s�b >>> new create excel�� & new create "xxx" �u�@��
                    xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLSX); //�w�q excel�榡�� XLSX
                    xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name);
                }

                
                driver.Navigate().GoToUrl(test_url); // �}�ҫ��w����
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                //driver.Manage().Window.Maximize(); // ���ù�����

                IWebElement go_to_redeem = driver.FindElement(By.XPath("//*[@id='layout_m_0_content_m_3_tab_0_HlkMeleageForm']"));
                go_to_redeem.Click(); // �I"�e���I��"
                System.Threading.Thread.Sleep(100);

                WebDriverWait redeem_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                redeem_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id='login']"))); //���ݪ���ݨ� �|���n�J dialog

                IWebElement creditcard_friend_login = driver.FindElement(By.XPath("//*[@id='layout_m_0_content_m_2_HlkCardLogin']/img"));
                creditcard_friend_login.Click();// �I"�H�Υd�͵n�J"
                System.Threading.Thread.Sleep(100);

                using (var reader = new StreamReader(UserDataList.csvpath)) //ŪCSV��
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<UserDataList>().ToList();
                    int k = 1;
                    foreach (var user in records)
                    {
                        if(k == 6)
                        {
                            IWebElement ID = driver.FindElement(By.XPath("//*[@id='CHID_T']"));
                            IWebElement CreditCard = driver.FindElement(By.XPath("//*[@id='cardNO_T']"));
                            IWebElement Expire_Mon = driver.FindElement(By.XPath("//*[@id='endM']"));
                            IWebElement Expire_Year = driver.FindElement(By.XPath("//*[@id='endY']"));
                            IWebElement CheckCode= driver.FindElement(By.XPath("//*[@id='cardCVV_T']"));
                            IWebElement Submit_button= driver.FindElement(By.XPath("//*[@id='confirm_btn']"));

                            ID.Clear();
                            ID.SendKeys(user.ID); // ��J�����Ҧr��
                            System.Threading.Thread.Sleep(100);

                            CreditCard.Clear();
                            CreditCard.SendKeys(user.CARDID_FULL); // ��J�d��
                            System.Threading.Thread.Sleep(100);

                            Expire_Mon.Clear();
                            Expire_Mon.SendKeys(user.MON); // ��J���Ĥ��
                            System.Threading.Thread.Sleep(100);

                            Expire_Year.Clear();
                            Expire_Year.SendKeys(user.YEAR); // ��J���Ħ~��
                            System.Threading.Thread.Sleep(100);

                            CheckCode.Clear();
                            CheckCode.SendKeys(user.CHECKCODE); //��J���ҽX
                            System.Threading.Thread.Sleep(100);

                            Submit_button.Click(); // �e�X
                            System.Threading.Thread.Sleep(100);

                            break;
                        }
                        k++;
                    }
                }

                IWebElement redeem_again = driver.FindElement(By.XPath("")); 
                redeem_again.Click(); // �e���I��
                System.Threading.Thread.Sleep(100);
                IWebElement submit_redeem_setting = driver.FindElement(By.XPath(""));
                submit_redeem_setting.Click(); //�I�e�X

                Tools.TakeScreenShot(@"d:\��Ũ��{�I��_" + testcase_name + "_" + browserType + "_" + Version + "_" + timesavepath + ".png", driver);
                System.Threading.Thread.Sleep(500);


                //�ˬd�I�ƿ�J���
                IWebElement airline_input_point = driver.FindElement(By.XPath(""));
                string[] array_asia_point = new string[] {"������", "����", "��", "@", "G", "0", "6666", "1", "20", "50", "999", "����r", "@#$*", "�I�I", "1a4", ""};
                
                //�w�qexcel���
                xlsSheet["A1"].Value = "�ˬd " + testcase_name + " '�I�� & ����' ���";
                xlsSheet["B1"].Value = "��ܰT��";
                xlsSheet["C1"].Value = "���G";
                xlsSheet["A1:C1"].Style.BottomBorder.SetColor("#ff6600"); // ���u"����"
                xlsSheet["A1:C1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double; //�[�����u


                int i = 3;
                foreach (var keyin_point in array_asia_point) // �]�j��H�v����J���
                {
                    string check_position = "A" + i;
                    string show_msg = "B" + i;
                    string test_result = "C" + i;

                    airline_input_point.Clear();
                    airline_input_point.SendKeys(keyin_point);

                    string total_point = driver.FindElement(By.Id("total")).GetAttribute("value");
                    string to_mile = driver.FindElement(By.Id("mile")).GetAttribute("value");
                    string display_keyin_point = airline_input_point.GetAttribute("value");
                    string unit_point = driver.FindElement(By.Id("")).GetAttribute("value");

                    submit_redeem_setting.Click();
                    string count_error = driver.FindElement(By.Id("counter-error")).Text;

                    xlsSheet[check_position].Value = "��J " + keyin_point;
                    xlsSheet[show_msg].Value = count_error;
                    xlsSheet[test_result].Value = " ==> �ڭn�� " + unit_point + " �I x " + display_keyin_point + " = " + total_point + " �I��" + testcase_name + " " + to_mile + " ��";
                    System.Threading.Thread.Sleep(100);
                    i++;
                }

                if (File.Exists(excel_savepath) != true)
                {
                    xlsWorkbook.SaveAs(excel_savepath);
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
