using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace Planirovanie
{
    public static class Waiting
    {

        public static void WaitForAjax(FirefoxDriver _firefox)
        {
            while (true) // Handle timeout somewhere
            {
                var ajaxIsComplete = (bool)(_firefox as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete)
                {
                    Thread.Sleep(500);
                    break;
                }
                Thread.Sleep(500);
            }
        }

        public static void WaitForAjaxWithoutSleep(FirefoxDriver _firefox)
        {
            while (true) // Handle timeout somewhere
            {
                var ajaxIsComplete = (bool)(_firefox as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete)
                {
                    break;
                }
            }
        }

        public static void WaitForTextInTitleAttribute(string locator, string text, FirefoxDriver _firefox)
        {

            const int waitRetryDelayMs = 1000; //шаг итерации (задержка)
            const int timeOut = 500; //время тайм маута 
            bool first = true;

            for (int milliSecond = 0; ; milliSecond += waitRetryDelayMs)
            {
                WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));

                try
                {
                    if (milliSecond > timeOut * 10000)
                    {
                        Debug.WriteLine("Timeout: Text " + text + " is not found ");
                        break; //если время ожидания закончилось (элемент за выделенное время не был найден)
                    }

                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                    if (_firefox.FindElement(By.XPath(locator)).GetAttribute("title") == text)
                    {
                        Thread.Sleep(2000);
                        if (!first) Debug.WriteLine("Text is found: " + text);
                        break; //если элемент найден
                    }

                    if (first) Debug.WriteLine("Waiting for text is present: " + text);

                    first = false;
                    Thread.Sleep(waitRetryDelayMs);
                }
                catch (Exception a)
                {
                    if (milliSecond < timeOut * 10000)
                        continue;
                    else
                    {
                        throw a;
                    }
                }
            }

        }

        public static void WaitPatternPresentInText(string locator, string text, FirefoxDriver _firefox)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            const int waitRetryDelayMs = 1000; //шаг итерации (задержка)
            const int timeOut = 500; //время тайм маута 

            for (int milliSecond = 0; ; milliSecond += waitRetryDelayMs)
            {
                if (milliSecond > timeOut * 10000)
                {
                    break; //если время ожидания закончилось (элемент за выделенное время не был найден)
                }

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                if (_firefox.FindElement(By.XPath(locator)).Text.ContainsIgnoreCase(text))
                {
                    Thread.Sleep(1000);
                    break; //если элемент найден
                }
                Thread.Sleep(waitRetryDelayMs);
            }
        }

        public static void WaitPatternPresentInAttribute(string locator, string attr, string text, FirefoxDriver _firefox)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            const int waitRetryDelayMs = 500; //шаг итерации (задержка)
            const int timeOut = 500; //время тайм маута 

            for (int milliSecond = 0; ; milliSecond += waitRetryDelayMs)
            {
                if (milliSecond > timeOut * 10000)
                {
                    break; //если время ожидания закончилось (элемент за выделенное время не был найден)
                }

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                if (_firefox.FindElement(By.XPath(locator)).GetAttribute(attr).ContainsIgnoreCase(text))
                {
                    break; //если элемент найден
                }
                Thread.Sleep(waitRetryDelayMs);
            }
        }
    }
}
