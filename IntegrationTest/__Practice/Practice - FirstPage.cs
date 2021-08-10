using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using Xunit.Abstractions;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium.Interactions;
using System.Web;

namespace AutomatedTest.IntegrationTest.__Practice
{
    public class �m��_�̷s�����P���i_List:IntegrationTestBase
    {
        public �m��_�̷s�����P���i_List(ITestOutputHelper output, Setup testsetup) : base(output, testsetup)
        {
            testurl = domain + "https://www.esunbank.com.tw/bank/personal#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void ���C�X�̷s�����P���i_ExportReport(string browser)
        {
            StartTestCase(browser, "�x������_�̷s�����P���i_���C���", "York");
            driver.FindElement(By.XPath("//*[@id='btnAntiFraud']")).Click(); // �����u������

            Tools.ScrollPageUpOrDown(driver, 300);

            INFO("�x�������̷s�����P���i");
            int newCounts = driver.FindElements(By.XPath("//*[@id='layout_0_maincontent_1_announce_0_divNotice']/ul/li")).Count; // �̷s����/���i�ƶq
            string newsAndAnnouncements =  driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_1_announce_0_divNotice']/ul")).Text; // �x������ "�̷s����/���i" list
            INFOfromCodeFormat($"{newsAndAnnouncements}");
           
            INFO("");
            INFO("�̷s�����P���i�s���T�{");
            for (int newindex = 0; newindex < newCounts; newindex++)
            {
                IWebElement NewsAnnouncementsList = driver.FindElement(By.XPath($"//*[@id='layout_0_maincontent_1_announce_0_RepAnnocement_Li1_{newindex}']/a"));
                string newsData = NewsAnnouncementsList.Text; // "���i/�̷s����" ���D

                Actions action = new Actions(driver);
                action.KeyDown(Keys.Control).MoveToElement(NewsAnnouncementsList).Click().Perform(); // �t�}�s�� "Ctrl+�ƹ�����"
                Thread.Sleep(300);
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // ������s�}�Ҫ�����
                Thread.Sleep(100);
                bool urlContain = driver.Url.Contains("https://www.esunbank.com.tw/bank/about/announcement/public/"); // �P�_��U���}�O�_�]�thttps://...

                if (urlContain == true) // (�N��page�b���i�P�n����)
                {
                    INFO($"�����P���i_{newindex + 1}: {newsData}_�e�����T��");
                    string newsArray = "";
                    for (int index = 0; index <= 4; index++) // �C�X�e�������i
                    {
                        string newsList = driver.FindElement(By.XPath($"//*[@id='layout_0_rightcontent_0_LsvAnnouncement_hlkLinkButton_{index}']")).Text;
                        newsArray = newsArray + newsList + "\r\n";
                    }
                    INFOfromCodeFormat($"{newsArray}");
                    driver.Close(); // ������U����
                    driver.SwitchTo().Window(driver.WindowHandles.First()); // ���^��_�l����
                }
                else
                {
                    INFO("");
                    INFO($"�����P���i_{newindex+1}: {newsData}");
                    string title = driver.FindElement(By.XPath("//*[@id='mainform']/div[10]/div[2]/div[1]/span")).Text ;
                    string content = driver.FindElement(By.XPath("//*[@id='mainform']/div[10]/div[2]/div[2]")).Text;
                    string stringArray = title + "\r\n\n" + content;
                    INFOfromCodeFormat($"{stringArray}");
                    driver.Close();
                    driver.SwitchTo().Window(driver.WindowHandles.First()); // ���^��_�l����
                }
            }

            CloseBrowser();
        }
    }
}


