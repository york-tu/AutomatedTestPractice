using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Xunit;
using System;
using System.Linq;

using System.IO;
using CSVHeader;
using CsvHelper;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;


namespace UnitTest1
{
    public class UnitTest1:TOOL
    {
        readonly string test_url = "https://www.esunbank.com.tw/bank/about/services/customer/message-board";
        private readonly string Version = "PC";
        public string _name ="";
        public string _id ="";
        public string _phone ="";
        public string _email ="";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void NAvigateToDpApp(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);


            using var driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                driver.Manage().Window.Maximize();


                for (int i = 2; i <= 22; i++) // initial i =2
                {
                    string timepath = System.DateTime.Now.ToString("yyyyMMddhhmmss");
                    int n = i - 1; // �� n �ӿ���

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 1, 35, 32, 11, 3, 3, 3, 14, 2, 1, 1, 2, 1, 10, 14, 2, 1, 1, 1, 1, 15 }; //�������������
                    while (j <= arrray[i - 2])
                    {
                        if (i == 22 && j == 14)
                        {
                            driver.Navigate().Refresh();
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000);
                        }
                        // �w�q���XPath
                        String Country_DropDownList = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span"; // �����U�Կ��XPath
                        String BranchName_DropDownList = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span"; // ����U�Կ��XPath
                        String Cuntry_Xpath = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[" + i + "]/span";
                        String BranchName_Xpath = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[" + j + "]/span";

                        // �Ҧ����w��
                        IWebElement FullNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[1]/td[2]/input"));
                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='citizenId']"));
                        IWebElement TelephoneColumn = driver.FindElement(By.XPath("//*[@id='phone']"));
                        IWebElement EMailColumn = driver.FindElement(By.XPath("//*[@id='email']"));
                        IWebElement BusinessItem_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/span"));
                        IWebElement CreditCardBusiness = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/ul/li[2]/span"));
                        IWebElement MessageArea = driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_2_comments']"));
                        IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]/a"));
                        IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));

                        
                        using (var reader = new StreamReader(@"C:\Users\axn01\source\repos\XUnitAutoTest\UserData\UserInfo.csv"))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            
                            //csv.Configuration.HasHeaderRecord = false;

                            var records = csv.GetRecords<UserDataList>().ToList();
                            int k = 1;
                            foreach (var user in records)
                            {
                                if (k == 2) // �� CSV �� k ����
                                {
                                     _name = user.NAME;
                                    _id = user.ID;
                                    _phone = user.PHONE;
                                    _email = user.EMAIL;


                                    break;
                                }
                                k++;
                            }

                            FullNameColumn.Clear();
                            FullNameColumn.SendKeys(_name); //��m�W
                            System.Threading.Thread.Sleep(100);

                            IdentityCardColumn.Clear();
                            IdentityCardColumn.SendKeys(_id); //�񨭤���
                            System.Threading.Thread.Sleep(100);

                            TelephoneColumn.Clear();
                            TelephoneColumn.SendKeys(_phone); //��s���q��
                            System.Threading.Thread.Sleep(100);

                            EMailColumn.Clear();
                            EMailColumn.SendKeys(_email); //�� E-MAIL

                            Find_Element(driver, EMailColumn);
                            driver.FindElement(By.XPath(Country_DropDownList)).Click(); // �i�} "�п�ܿ���" �U�Կ��
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(Cuntry_Xpath)).Click(); // �I��@�� "����" 
                            System.Threading.Thread.Sleep(300);
                            System.Diagnostics.Debug.WriteLine("2. �I����� " + n + " �ӿ���");
                            System.Threading.Thread.Sleep(300);
                            


                            Find_Element(driver, EMailColumn);
                            driver.FindElement(By.XPath(BranchName_DropDownList)).Click(); // �i�} "����" �U�Կ��
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(BranchName_Xpath)).Click(); // �I��@�� "����"
                            System.Threading.Thread.Sleep(300);
                            System.Diagnostics.Debug.WriteLine("4. �I����� " + j + " �Ӥ���");
                            System.Threading.Thread.Sleep(300);

                           

                            Find_Element(driver, EMailColumn);
                            System.Threading.Thread.Sleep(100);
                            BusinessItem_DropDownList.Click(); //�I�d���~�ȤU�Կ��
                            System.Threading.Thread.Sleep(300);
                            CreditCardBusiness.Click(); //��H�Υd�~��
                            System.Threading.Thread.Sleep(300);

                            MessageArea.Clear();
                            MessageArea.SendKeys(
                                "[Automation Test on " + browserType + " in " + Version + " ]" + "\r\n"
                                + "��� on �� " + n + " �ӿ����� " + j + " �Ӥ���" + "\r\n"
                                //+ "=================================================" + "\r\n"
                                //+ "Total 21 ��, 154 ����" + "\r\n"
                                //+ "����(�a��index)(����ƶq):" + "\r\n"
                                //+ "�򶩥�(1)(1),      �O�_��(2)(35),      �s�_��(3)(32),      ��饫(4)(11),      �s�˥�(5)(3),      "
                                //+ "�s�˿�(6)(3),      �]�߿�(7)(3),        �O����(8)(14),      ���ƿ�(9)(2),        �n�뿤(10)(1)," + "\r\n"
                                //+ "���L��(11)(1),    �Ÿq��(12)(2),      �Ÿq��(13)(1),      �x�n��(14)(10),    ������(15)(14),      "
                                //+ "�̪F��(16)(2),    �y����(17)(1),      �Ὤ��(18)(1),      �O�F��(19)(1),      ���(20)(1),       ���~�a��(21)(15)."
                                ); // �� "�d�����e"

                            TakeScreenShot("d:\\�� " + n + " �ӿ����� " + j + " �Ӥ��� snapshot_" + browserType + "_" + timepath +".png", driver); // snapshot��U�e��
                            System.Threading.Thread.Sleep(1000);
                           // ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='submit']")), "d:\\" + browserType + timepath +".png");
                            System.Threading.Thread.Sleep(3000);

                            //IHaveReadRadioButtin.Click(); // �I "�ڤw�\Ū" radio button
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("13. �Ŀ�'�ڤw�\Ū'");

                            SCrollToElement(driver, MessageArea);
                            System.Threading.Thread.Sleep(3000);
                            //SubmitButton.Click(); // �I "�e�X" ���s
                            System.Threading.Thread.Sleep(300);
                            //ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]")), "d:\\" + browserType + timepath + ".png");
                            // WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                            // wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // ���ݪ���ݨ�q������ 

                            // driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // �I�q������ "X" ���s
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("15. �����q������");

                            // WebDriverWait wait_to_see_submit_button = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                            // wait_to_see_submit_button.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id='submit']"))); // ���ݪ���ݨ�"�d����"
                            // System.Threading.Thread.Sleep(100);

                            driver.FindElement(By.XPath("//*[@id='scrollUp']")).Click(); //�e���^�쳻��
                            System.Threading.Thread.Sleep(500);
                            System.Diagnostics.Debug.WriteLine("16. �^�������");
                            System.Threading.Thread.Sleep(100);

                            j++;
                        }
                    }
                    driver.Quit();
                }
            }



        }
    }
}
