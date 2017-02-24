using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

namespace Planirovanie.CheckStadaPlan
{
    public class PageElementsAdditional
    {
        private readonly IWebDriver _driver;

        public PageElementsAdditional(IWebDriver firefox)
        {
            _driver = firefox;

        }

        public const string ClosePreparationListButton = "html/body/div[4]/div[3]/div/button[1]";
        public const string TopMenuOdobreniePlanovButton = ".//*[@id='tabs']/ul/li[6]/a";
        public const string TableOdobrenie = ".//*[@id='dep_info']";
        public const string TableOdobrenieRows = ".//*[@id='dep_info']/tbody/tr";
        public const string TopMenuPlanyPoTerritoriamButton = ".//*[@id='tabs']/ul/li[4]/a";

    }
}
