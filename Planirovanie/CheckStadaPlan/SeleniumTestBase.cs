using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace Planirovanie.CheckStadaPlan
{
    [TestFixture]
    public abstract class SeleniumTestBase
    {
        protected IWebDriver Driver;
       
        public WebDriverWait Wait
        {
            get { return new WebDriverWait(Driver, TimeSpan.FromSeconds(120)); }
        }

        public PageElements PageElements
        {
            get { return new PageElements(Driver as FirefoxDriver);}
        }

        public PageElementsAdditional PageElementsAdditional
        
        {
            get { return new PageElementsAdditional(Driver); }
        }

       /* public Methods Methods

        {
            get { return new Methods(Driver as FirefoxDriver); }
        }*/

        [TestFixtureSetUp]
        public void TestInitialize()
        {
            Driver = new FirefoxDriver();
           }

        [TestFixtureTearDown]
        public void TestCleanup()
        {
            Driver.Quit();
        }

        [SetUp]
        public void OneSetUp()
        {
            Driver.Manage().Window.Maximize();
           
        }


        [TearDown]
        public void OneTearDown()
        {
           /* Driver.Manage().Cookies.DeleteAllCookies();*/
        }

    }
}
