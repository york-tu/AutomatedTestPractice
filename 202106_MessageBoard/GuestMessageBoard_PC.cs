using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using Utilities;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Globalization;


namespace GuestMessageBoardTest
{
    public class �X�ȯd����������Ʒ��վ�_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/about/services/customer/message-board";
        private readonly string Version = "PC";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                driver.Manage().Window.Maximize();


                for (int i = 22; i <= 22; i++) // initial i =2
                {

                    int n = i - 1; // �� n �ӿ���

                    int j = 10; // initial j = 1
                    int[] arrray = new int[] { 1, 35, 32, 11, 3, 3, 3, 14, 2, 1, 1, 2, 1, 10, 14, 2, 1, 1, 1, 1, 15 }; //�������������
                    while (j <= arrray[i - 2])
                    {
                        if (i == 22 && j == 14)
                        {
                            driver.Navigate().Refresh();
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000);
                        }

                        // �w�q���XPath
                        string Country_DropDownList = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span"; // �����U�Կ��XPath
                        string BranchName_DropDownList = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span"; // ����U�Կ��XPath
                        string Cuntry_Xpath = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[" + i + "]/span";
                        string BranchName_Xpath = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[" + j + "]/span";

                        string time = System.DateTime.Now.ToString("yyyyMMddHHmmss");

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

                        string csvpath = $@"{UserDataList.folderpath}\testdata\UserInfo.csv";
                        using (var reader = new StreamReader(csvpath))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            var records = csv.GetRecords<UserDataList>().ToList();
                            int k = 1;
                            foreach (var user in records)
                            {
                                if (k == 2) // �� CSV �� k ����
                                {
                                    FullNameColumn.Clear();
                                    FullNameColumn.SendKeys(user.NAME); //��m�W
                                    System.Threading.Thread.Sleep(100);

                                    IdentityCardColumn.Clear();
                                    IdentityCardColumn.SendKeys(user.ID); //�񨭤���
                                    System.Threading.Thread.Sleep(100);

                                    TelephoneColumn.Clear();
                                    TelephoneColumn.SendKeys(user.PHONE); //��s���q��
                                    System.Threading.Thread.Sleep(100);

                                    EMailColumn.Clear();
                                    EMailColumn.SendKeys(user.EMAIL); //�� E-MAIL
                                    System.Threading.Thread.Sleep(100);


                                    break;
                                }
                                k++;
                            }

                            
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(Country_DropDownList)).Click(); // �i�} "�п�ܿ���" �U�Կ��
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(Cuntry_Xpath)).Click(); // �I��@�� "����" 
                            System.Threading.Thread.Sleep(300);

                            driver.FindElement(By.XPath(BranchName_DropDownList)).Click(); // �i�} "����" �U�Կ��
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(BranchName_Xpath)).Click(); // �I��@�� "����"
                            System.Threading.Thread.Sleep(300);

                            string snapshotpath = $@"{UserDataList.folderpath}\SnapshotFolder\GuestMessageBoard";
                            Tools.CreateSnapshotFolder(snapshotpath);
                            System.Threading.Thread.Sleep(100);

                            Tools.ElementTakeScreenShot(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]")),
                                $@"{snapshotpath}\�� {n} �ӿ����� {j} �Ӥ���_��� snapshot_{time}.png"); //�����I��

              
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

                            Tools.SCrollToElement(driver, TelephoneColumn);
                            Tools.TakeScreenShot($@"d:\�� {n} �ӿ����� {j} �Ӥ��� snapshot_{browserType}_ {time}.png", driver); // snapshot��U�e��
                            System.Threading.Thread.Sleep(500);




                            // IHaveReadRadioButtin.Click(); // �I "�ڤw�\Ū" radio button
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("13. �Ŀ�'�ڤw�\Ū'");

                            // SCrollToElement(driver, MessageArea);
                            // System.Threading.Thread.Sleep(3000);

                            // SubmitButton.Click(); // �I "�e�X" ���s
                            // System.Threading.Thread.Sleep(300);

                            // ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]")), "d:\\" + browserType + timepath + ".png");
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

                            j++;
                        }
                    }
                }
                driver.Quit();
            }
        }
    }
}
    

