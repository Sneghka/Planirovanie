using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace Planirovanie
{
    public static class Helper
    {

        private static readonly FirefoxDriver _firefox;

        public static bool IsElementPresent(By by)
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

        public static void TryToClickWithoutException(string locator, IWebElement element)
        {
            var MAX_STALE_ELEMENT_RETRIES = 100;
            var action = new Actions(_firefox);
            var retries = 0;
            while (true)
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait(new SystemClock(), _firefox, TimeSpan.FromSeconds(120), TimeSpan.FromSeconds(5));
                    wait.Until(ExpectedConditions.ElementExists(By.XPath(locator)));

                    _firefox.FindElement(By.XPath(locator)).Click();
                    Waiting.WaitForAjax();
                    return;
                }
                catch (Exception e)
                {
                    if (retries < MAX_STALE_ELEMENT_RETRIES)
                    {
                        retries++;
                        Debug.WriteLine(retries);
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
