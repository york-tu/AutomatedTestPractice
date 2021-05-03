using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;

namespace System
{
    public class TOOL
    {
        public static void Find_Element(IWebDriver driver, IWebElement element) // 畫面定位
        {
            Actions action = new(driver);
            action.MoveToElement(element).Perform();

        }
        public static void TakeScreenShot(String savepath, IWebDriver driver) // 截圖 
        {
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(savepath, ScreenshotImageFormat.Png);
        }
    }
}
