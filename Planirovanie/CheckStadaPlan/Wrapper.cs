using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace Planirovanie.CheckStadaPlan
{
    public static class Wrapper
    {
        private static string _domain;

        public static string GetUrl(string url)
        {
            _domain = url;
            return _domain;
        }

        public static void TryToClickWithoutException(string locator, IWebDriver driver, WebDriverWait wait)
        {
            var MAX_STALE_ELEMENT_RETRIES = 100;
            var retries = 0;
            while (true)
            {
                try
                {
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                    driver.FindElement(By.XPath(locator)).Click();
                    Waiting.WaitForAjax(driver as FirefoxDriver);
                    return;
                }
                catch (Exception e)
                {
                    if (retries < MAX_STALE_ELEMENT_RETRIES)
                    {
                        Thread.Sleep(100);
                        retries++;
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
