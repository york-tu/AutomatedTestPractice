using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.Drawing;

namespace References
{
    public class Tools
    {

        public static void Find_Element(IWebDriver driver, IWebElement element) // 畫面定位 (selenium)
        {
            Actions action = new Actions(driver);
            action.MoveToElement(element).Perform();

        }
        public static void TakeScreenShot(string savepath, IWebDriver driver) // 截圖當下畫面 (selenium)
        {
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(savepath, ScreenshotImageFormat.Png);
        }

        public static void ElementTakeScreenShot(IWebElement webElement, string savepath) // Only截元件圖 (selenium)
        {
            var elementScreenshot = (webElement as ITakesScreenshot).GetScreenshot();
            elementScreenshot.SaveAsFile(savepath);
        } 

        public static void SnapshotFullScreen (string savepath) //全螢幕截圖 (C#)
        {
            Bitmap myimage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myimage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0),new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            myimage.Save(savepath);
        }
        public static void SCrollToElement(IWebDriver driver, IWebElement element) // 畫面滾到target位置
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", element); // Viewport對頂部對齊
            Actions act = new Actions(driver);
            act.MoveToElement(element).Perform();

        }


    }
}
