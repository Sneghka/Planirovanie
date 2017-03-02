using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;



namespace Planirovanie.CheckStadaPlan
{
    [TestFixture]
    public class CheckDataBeforeSending : SeleniumTestBase
    {
        const string test = "http://test.stada.bi.morion.ua/";
        const string logoutTest = "http://test.stada.bi.morion.ua/logout.php";
        const string dev = "http://dev.stada.bi.morion.ua/";
        const string logoutDev = "http://dev.stada.bi.morion.ua/logout.php";
        const string stada = "http://stada.bi.morion.ua";
        const string logoutStada = "http://stada.bi.morion.ua/logout.php";

        #region USAGE EXAMPLES

        /* [Test, Timeout(10000)]
         public void Test1()
         {
             Driver.Navigate().GoToUrl(Wrapper.GetUrl("https://www.yandex.ua/"));
             Assert.IsTrue(Driver.Title == "Яндекс");
         }


         [Test, Timeout(10000)]
         public void Test2()
         {
             Driver.Navigate().GoToUrl(Wrapper.GetUrl(test));
             Assert.IsTrue(Driver.Title == "Планировщик продаж - Авторизация (DEV)");
         }

         [Test, Timeout(10000)]
         public void Test3()
         {
             Driver.Navigate().GoToUrl(Wrapper.GetUrl(stada));
             Assert.IsTrue(Driver.Title == "Планировщик продаж - Авторизация (DEV)");
         }
 */
        #endregion

        [Test]
        public void CheckPlans()
        {
            var startTime = DateTime.Now;
            var methods = new Methods(Driver as FirefoxDriver);
            var firefox2 = new FirefoxDriver();

            var months = 3;
            //Перед считыванием вручную разбить задвоеные БЮ (напр 31,94 - копипастом продублировать всё для 31,а потом тоже самое для 94)
            methods.StoreExcelDataTerritoriiSpravochnik("zone_of_resp"); // считываем закладку Зоны ответственности из справочника

            Console.WriteLine("Считали зоны ответственности");
            methods.StoreExcelDataEmailSpravochik("email"); // считываем закладку Email из справочника
            Console.WriteLine("Считали email");
            //Methods.StoreLoginPasswordFromExcel();
            methods.LoginStada(test, "user_1340", "m600e");

            methods.GoToOdobreniePlanovTab();
            // Methods.CalculateAllPlans(stada); // ЗАГЛУШКА
            methods.ReadPlanFor33BU(test, months, firefox2, logoutTest);
            Console.WriteLine("Считали BU33");
            methods.CheckCalculatedPlans(test, months, firefox2, logoutTest);


            var finishTime = DateTime.Now;
            var lasting = startTime - finishTime;
            Console.WriteLine("");
            Console.WriteLine("ВРЕМЯ ПРОВЕРКИ ПЛАНОВ  - " + lasting);
            firefox2.Quit();
            Driver.Quit();

        }






    }
}
