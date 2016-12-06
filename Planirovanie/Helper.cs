using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace Planirovanie
{
    public static class Helper
    {

       /* private static readonly FirefoxDriver _firefox;*/

        public static bool IsElementPresent(By by, FirefoxDriver _firefox)
        {
            try
            {
                _firefox.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public static void TryToClickWithoutException(string locator, FirefoxDriver firefox)
        {
           
            var MAX_STALE_ELEMENT_RETRIES = 100;
            var retries = 0;
            while (true)
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait(new SystemClock(), firefox, TimeSpan.FromSeconds(10), TimeSpan.FromSeconds(5));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                    firefox.FindElement(By.XPath(locator)).Click();
                    Waiting.WaitForAjax(firefox);
                    return;
                }
                catch (Exception e)
                {
                    if (retries < MAX_STALE_ELEMENT_RETRIES)
                    {
                        retries++;
                        Debug.WriteLine("Try to click - " + retries);
                        continue;
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
        }
    }
}
