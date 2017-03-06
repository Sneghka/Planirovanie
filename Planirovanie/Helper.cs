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
            int i = 0;
            while (i < 5)
            {
                try
                {
                    _firefox.FindElement(by);
                    return true;
                }
                catch (NoSuchElementException)
                {
                   i++;
                    Thread.Sleep(500);
                }
            }
            return false;
        }

        public static void TryToClickWithoutException(string locator, FirefoxDriver firefox)
        {

            var MAX_STALE_ELEMENT_RETRIES = 100;
            var retries = 0;
            while (true)
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait(new SystemClock(), firefox, TimeSpan.FromSeconds(10),
                        TimeSpan.FromSeconds(5));
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

        public static void CompareIdLists(List<int> listSpravochnik, List<int> listPlanirovschik)
        {

            var lengthFile1 = listSpravochnik.Count;
            var lengthFile2 = listPlanirovschik.Count;

            int x = 0;
            int y = 0;

            while (x < lengthFile1 || y < lengthFile2)
            {
                if ((x < lengthFile1 && listSpravochnik[x] < 0) || (y < lengthFile2 && listPlanirovschik[y] < 0))
                {
                    Console.WriteLine("Id территории отрицательное - сверка не прошла (" + listPlanirovschik[y] + ")");
                    return;
                }
                if (x < lengthFile1 && y < lengthFile2 && listSpravochnik[x] < listPlanirovschik[y])
                {
                    Console.WriteLine("Территория из Справочника - ID " + listSpravochnik[x] +
                                      " -  не содержится в Планировщике");
                    x++;
                    continue;
                }
                if (x < lengthFile1 && y < lengthFile2 && listSpravochnik[x] > listPlanirovschik[y])
                {
                    Console.WriteLine("Территория из Планировщика - ID " + listPlanirovschik[y] +
                                      " -  не содержится в Справочнике");
                    y++;
                    continue;
                }
                if (x < lengthFile1 && y < lengthFile2 && listSpravochnik[x] == listPlanirovschik[y])
                {
                    x++;
                    y++;
                    continue;
                }

                if (y >= lengthFile2)
                {
                    Console.WriteLine("Территория из Справочника - ID " + listSpravochnik[x] +
                                      " -  не содержится в Планировщике");
                    x++;
                    continue;
                }
                if (x >= lengthFile1)
                {
                    Console.WriteLine("Территория из Планировщика - ID " + listPlanirovschik[y] +
                                     " -  не содержится в Справочнике");
                    y++;
                }
            }
        }
    }
}
