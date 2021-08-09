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
    public class 練習_最新消息與公告_List:IntegrationTestBase
    {
        public 練習_最新消息與公告_List(ITestOutputHelper output, Setup testsetup) : base(output, testsetup)
        {
            testurl = domain + "https://www.esunbank.com.tw/bank/personal#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 條列出最新消息與公告_ExportReport(string browser)
        {
            StartTestCase(browser, "官網首頁_最新消息與公告_條列資料", "York");
            driver.FindElement(By.XPath("//*[@id='btnAntiFraud']")).Click(); // 關掉彈跳視窗

            Tools.ScrollPageUpOrDown(driver, 300);

            INFO("官網首頁最新消息與公告");
            int newCounts = driver.FindElements(By.XPath("//*[@id='layout_0_maincontent_1_announce_0_divNotice']/ul/li")).Count; // 最新消息/公告數量
            string newsAndAnnouncements =  driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_1_announce_0_divNotice']/ul")).Text; // 官網首頁 "最新消息/公告" list
            INFOfromCodeFormat($"{newsAndAnnouncements}");
           
            INFO("");
            INFO("最新消息與公告連結確認");
            for (int newindex = 0; newindex < newCounts; newindex++)
            {
                IWebElement NewsAnnouncementsList = driver.FindElement(By.XPath($"//*[@id='layout_0_maincontent_1_announce_0_RepAnnocement_Li1_{newindex}']/a"));
                string newsData = NewsAnnouncementsList.Text; // "公告/最新消息" 標題

                Actions action = new Actions(driver);
                action.KeyDown(Keys.Control).MoveToElement(NewsAnnouncementsList).Click().Perform(); // 另開新頁 "Ctrl+滑鼠左鍵"
                Thread.Sleep(300);
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // 切換到新開啟的分頁
                Thread.Sleep(100);
                bool urlContain = driver.Url.Contains("https://www.esunbank.com.tw/bank/about/announcement/public/"); // 判斷當下網址是否包含https://...

                if (urlContain == true) // (代表page在公告與聲明頁)
                {
                    INFO($"消息與公告_{newindex + 1}: {newsData}_前五筆訊息");
                    string newsArray = "";
                    for (int index = 0; index <= 4; index++) // 列出前五筆公告
                    {
                        string newsList = driver.FindElement(By.XPath($"//*[@id='layout_0_rightcontent_0_LsvAnnouncement_hlkLinkButton_{index}']")).Text;
                        newsArray = newsArray + newsList + "\r\n";
                    }
                    INFOfromCodeFormat($"{newsArray}");
                    driver.Close(); // 關掉當下分頁
                    driver.SwitchTo().Window(driver.WindowHandles.First()); // 切回到起始分頁
                }
                else
                {
                    INFO("");
                    INFO($"消息與公告_{newindex+1}: {newsData}");
                    string title = driver.FindElement(By.XPath("//*[@id='mainform']/div[10]/div[2]/div[1]/span")).Text ;
                    string content = driver.FindElement(By.XPath("//*[@id='mainform']/div[10]/div[2]/div[2]")).Text;
                    string stringArray = title + "\r\n\n" + content;
                    INFOfromCodeFormat($"{stringArray}");
                    driver.Close();
                    driver.SwitchTo().Window(driver.WindowHandles.First()); // 切回到起始分頁
                }
            }

            CloseBrowser();
        }
    }
}


