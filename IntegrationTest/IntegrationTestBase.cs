using System;
using Xunit;
using Xunit.Abstractions;
using AventStack.ExtentReports;
using AutomatedTest.Utilities;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using AventStack.ExtentReports.MarkupUtils;

namespace AutomatedTest.IntegrationTest
{

    public class IntegrationTestBase:XunitContextBase, IClassFixture<Setup>
    {
        protected ExtentReportHelper _extentReport;
        protected ExtentTest _extentTObj;
        protected string testurl; // 測試用的網址
        protected string testFullName; // 測試案例全名
        protected string browserName; // 測試用的瀏覽器
        protected string runtimeEnvironment; // 執行環境
        protected bool result = false; // 初始測試結果
        protected string domain;
        protected bool getreport;
        private IConfiguration _config;
        protected ServiceProvider _serviceProvider;
        protected ITestOutputHelper _output;
        protected IWebDriver driver;
        BrowserHelper browser;

        public IntegrationTestBase(ITestOutputHelper output, Setup testSetup) : base(output)
        {
            _serviceProvider = testSetup.ServiceProvider;

            _config = _serviceProvider.GetService<IConfiguration>();
            runtimeEnvironment = _config.GetValue<string>("RuntimeEnvironment");
            string getjson = "Domain: " + runtimeEnvironment;

            //組成url
            domain= _config.GetValue<string>(getjson);

            //確認是否產生report
            getreport = Convert.ToBoolean(_config.GetValue<string>("GetReport"));
            testFullName = Context.UniqueTestName;

        }


        /// <summary>
        /// 開始測試 & 初始設定
        /// </summary>
        /// <param name="browserArg"></param>
        /// <param name="reportName"></param>
        /// <param name="testURL"></param>
        public void StartTestCase(string browserArg, string reportName, string tester)
        {
            browser = new BrowserHelper(browserArg);
            this.driver = browser.driver;
            driver.Navigate().GoToUrl(testurl);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
            driver.Manage().Window.Maximize();

            if (getreport)
            {
                CreateReport(reportName, tester);
                browserName = browserArg;
            }
        }
        public void CreateReport(string reportName, string tester)
        {
            _extentReport = new ExtentReportHelper(reportName, tester);
            _extentReport.Init();
            _extentReport.CreateTestCase(testFullName, "");
            _extentReport.GetTestEnvironment(runtimeEnvironment);
            _extentReport.TestCaseDriverVersion(browser.GetBrowserName() + " v" + browser.GetBrowserVersion());
            _extentTObj = _extentReport.ExtentTestObjects;
        }
        public void CloseBrowser()
        {
            this.driver.Quit();
            if (getreport)
            {
                _extentReport.ExportReport();
            }
        }
        public bool CloseWindows(string window)
        {
            try
            {
                this.driver.SwitchTo().Window(window);
                this.driver.Quit();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public void StartTestCaseForCustomizedSize(string browserArg, string reportName,string tester ,int position_X, int position_Y, int width, int height) // For 客製化開啟的視窗位置 + 大小
        {
            browser = new BrowserHelper(browserArg);
            this.driver = browser.driver;
            driver.Navigate().GoToUrl(testurl);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
            driver.Manage().Window.Position = new Point(position_X, position_Y); //設定網頁開啟在畫面什麼位置
            driver.Manage().Window.Size = new Size(width, height); // 設定開啟的網頁大小

            if (getreport)
            {
                CreateReport(reportName, tester);
                browserName = browserArg;
            }
        }


        ///<summary>
        ///Extent Report Logs
        ///</summary> 
        public void FAIL(string content)
        {
            if (getreport)
            {
                _extentTObj.Fail(MarkupHelper.CreateLabel(content, ExtentColor.Red));
            }
        }
        public void PASS(string content)
        {
            if (getreport)
            {
                _extentTObj.Pass(MarkupHelper.CreateLabel(content, ExtentColor.Green));
            }
        }
        public void ERROR(string content)
        {
            if (getreport)
            {
                _extentTObj.Error(MarkupHelper.CreateLabel(content, ExtentColor.Red));
            }
        }
        public void INFO(string content)
        {
            if (getreport)
            {
                _extentTObj.Info(content);

            }
        }
        public void WARNING(string content)
        {
            if (getreport)
            {
                _extentTObj.Warning(MarkupHelper.CreateLabel(content, ExtentColor.Orange));

            }
        }


        ///<summary>
        ///Tesseract 圖片辨識
        ///</summary> 
        Image img;
        Bitmap bimg;
        string Security_Code;
        private Image Getimage(string sourcepath) //讀取驗證碼圖片
        {
            Bitmap img = new Bitmap(sourcepath);
            return img;
        }
        public string executing() // 執行圖片處理
        {
            Image lookimg;

            //轉灰階
            convert2GrayScale();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());

            //去雜線
            RemoteNoiseLineByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());

            //去外框
            ClearPictureBorder(2);

            //去雜點
            RemoteNoisePointByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());

            //補點
            AddNoisePointByPixels();
            img = Image.FromHbitmap(bimg.GetHbitmap());
            parseCaptchaStr(img);

            return Security_Code;
        }
        public string TesseractOCRIdentify(string filepath, double imagescale) // 執行圖片辨識方法
        {
            img = Getimage(filepath);
            img = new Bitmap(img, (int)(img.Width * imagescale), (int)(img.Height * imagescale));
            bimg = new Bitmap(img);
            string words = executing();
            return words;
        }
        private void convert2GrayScale() // 轉灰階
        {
            for (int i = 0; i < bimg.Width; i++)
            {
                for (int j = 0; j < bimg.Height; j++)
                {
                    Color pixelColor = bimg.GetPixel(i, j);
                    byte r = pixelColor.R;
                    byte g = pixelColor.G;
                    byte b = pixelColor.B;

                    byte gray = (byte)(0.299 * (float)r + 0.587 * (float)g + 0.114 * (float)b);
                    r = g = b = gray;
                    pixelColor = Color.FromArgb(r, g, b);

                    bimg.SetPixel(i, j, pixelColor);
                }
            }
        }
        private void ClearPictureBorder(int pBorderWidth) //去邊框
        {
            for (int i = 0; i < bimg.Height; i++)
            {
                for (int j = 0; j < bimg.Width; j++)
                {
                    if (i < pBorderWidth || j < pBorderWidth || j > bimg.Width - 1 - pBorderWidth || i > bimg.Height - 1 - pBorderWidth)
                        bimg.SetPixel(j, i, Color.FromArgb(255, 255, 255));
                }
            }
        }
        private class NoisePoint
        {
            public int X { get; set; }
            public int Y { get; set; }
        }
        private void RemoteNoisePointByPixels() //去雜點
        {
            List<NoisePoint> points = new List<NoisePoint>();

            for (int k = 0; k < 5; k++)
            {
                for (int i = 0; i < bimg.Height; i++)
                    for (int j = 0; j < bimg.Width; j++)
                    {
                        int flag = 0;
                        int garyVal = 255;
                        // 檢查上相鄰像素
                        if (i - 1 > 0 && bimg.GetPixel(j, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && bimg.GetPixel(j, i + 1).R != garyVal) flag++;
                        if (j - 1 > 0 && bimg.GetPixel(j - 1, i).R != garyVal) flag++;
                        if (j + 1 < bimg.Width && bimg.GetPixel(j + 1, i).R != garyVal) flag++;
                        if (i - 1 > 0 && j - 1 > 0 && bimg.GetPixel(j - 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j - 1 > 0 && bimg.GetPixel(j - 1, i + 1).R != garyVal) flag++;
                        if (i - 1 > 0 && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i + 1).R != garyVal) flag++;

                        if (flag < 3)
                            points.Add(new NoisePoint() { X = j, Y = i });
                    }
                foreach (NoisePoint point in points)
                    bimg.SetPixel(point.X, point.Y, Color.FromArgb(255, 255, 255));
            }
        }
        private void RemoteNoiseLineByPixels() //去噪音線
        {
            for (int i = 0; i < bimg.Height; i++)
                for (int j = 0; j < bimg.Width; j++)
                {
                    int grayValue = bimg.GetPixel(j, i).R;
                    if (grayValue <= 255 && grayValue >= 160)
                        bimg.SetPixel(j, i, Color.FromArgb(255, 255, 255));
                }
        }
        private void AddNoisePointByPixels()  //補點
        {
            List<NoisePoint> points = new List<NoisePoint>();

            for (int k = 0; k < 1; k++)
            {
                for (int i = 0; i < bimg.Height; i++)
                    for (int j = 0; j < bimg.Width; j++)
                    {
                        int flag = 0;
                        int garyVal = 255;
                        // 檢查上相鄰像素
                        if (i - 1 > 0 && bimg.GetPixel(j, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && bimg.GetPixel(j, i + 1).R != garyVal) flag++;
                        if (j - 1 > 0 && bimg.GetPixel(j - 1, i).R != garyVal) flag++;
                        if (j + 1 < bimg.Width && bimg.GetPixel(j + 1, i).R != garyVal) flag++;
                        if (i - 1 > 0 && j - 1 > 0 && bimg.GetPixel(j - 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j - 1 > 0 && bimg.GetPixel(j - 1, i + 1).R != garyVal) flag++;
                        if (i - 1 > 0 && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i + 1).R != garyVal) flag++;

                        if (flag >= 7)
                            points.Add(new NoisePoint() { X = j, Y = i });
                    }
                foreach (NoisePoint point in points)
                    bimg.SetPixel(point.X, point.Y, Color.FromArgb(0, 0, 0));
            }
        }
        private void parseCaptchaStr(Image image) // 讀出字串
        {
            Ocr ocr = new Ocr();
            Bitmap BmpSource = new Bitmap(image);
            ocr.DoOCRMultiThred(BmpSource, "eng");
            Security_Code = ocr.GetSecurity_Code();
        }
    }

    public class Ocr
    {
        public string Security_Code;

        public void DumpResult(List<tessnet2.Word> result)
        {
            var sb = new StringBuilder();
            result.ForEach(c => {
                sb.Append(c.Text);
            });
            Security_Code = sb.ToString();
        }

        public string GetSecurity_Code()
        {
            return Security_Code;
        }

        public List<tessnet2.Word> DoOCRNormal(Bitmap image, string lang)
        {
            tessnet2.Tesseract ocr = new tessnet2.Tesseract();
            ocr.Init(null, lang, false);
            List<tessnet2.Word> result = ocr.DoOCR(image, Rectangle.Empty);
            DumpResult(result);
            return result;
        }

        ManualResetEvent m_event;

        public void DoOCRMultiThred(Bitmap image, string lang)
        {
            tessnet2.Tesseract ocr = new tessnet2.Tesseract();
            ocr.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789");
            var path = string.Concat(Application.StartupPath, @"\tessdata");
            ocr.Init(path, lang, false);
            // If the OcrDone delegate is not null then this'll be the multithreaded version
            ocr.OcrDone = new tessnet2.Tesseract.OcrDoneHandler(Finished);
            // For event to work, must use the multithreaded version
            m_event = new ManualResetEvent(false);
            ocr.DoOCR(image, Rectangle.Empty);
            // Wait here it's finished
            m_event.WaitOne();
        }

        public void Finished(List<tessnet2.Word> result)
        {
            DumpResult(result);
            m_event.Set();
        }
    }

}

