using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Drawing;
using java.awt.image;
using javax.imageio;

namespace System
{
    public class TOOL
    {

        public static void Find_Element(IWebDriver driver, IWebElement element) // 畫面定位
        {
            Actions action = new Actions(driver);
            action.MoveToElement(element).Perform();

        }
        public static void TakeScreenShot(String savepath, IWebDriver driver) // 截圖 
        {
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(savepath, ScreenshotImageFormat.Png);
        }

        public static void ElementTakeScreenShot(IWebElement webElement, string savepath)
        {
            var elementScreenshot = (webElement as ITakesScreenshot).GetScreenshot();
            elementScreenshot.SaveAsFile(savepath);
        }

        public static void SCrollToElement(IWebDriver driver, IWebElement element)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", element); // Viewport對頂部對齊
            Actions act = new Actions(driver);
            act.MoveToElement(element).Perform();

        }


    }
}
