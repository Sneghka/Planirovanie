using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium.Firefox;

namespace Planirovanie
{
    [TestFixture]
    public class CheckPlanirovschikFitoFharm
    {

        const string fitofharm = "http://test.fito.bi.morion.ua/";
        const string lpgoutFitoFarm = "http://test.fito.bi.morion.ua/logout.php";


        [Test]
        public void CheckPreparationsName()
        {
            var firefox = new FirefoxDriver();
            var fitoFharm = new FitoFharmMethods(firefox);

            fitoFharm.StoreExcelDataAny(@"D:\Sneghka\Selenium\Projects\Planirovschik\FitoPharm.xlsx");
            Console.WriteLine("Excel was stored");
            fitoFharm.LoginPlanirovschik(fitofharm, "pm", "fito");
            fitoFharm.StorePreparationNamesFromPlanirovschik();
            Console.WriteLine("Planirovschik was stored");
            Console.WriteLine("Сравниваем ексель с планировщиком");
            fitoFharm.CompareExcelWithWeb();
            Console.WriteLine("Сравниваем Планировщик с екселем");
            fitoFharm.CompareWebWithExcel();

            firefox.Quit();
        }

        [Test]
        public void ClickChainsAccept()
        {
            var firefox = new FirefoxDriver();
           var fitoFharmMethods = new FitoFharmMethods(firefox);

            string[] fitoFharmPM = { "pm" };
            string[] fitoFharmOTC = { "ua_otc", "center_otc", "south_otc", "east_otc", "south_east_otc", "west_otc", "nord_east_otc", "center_west_otc" };
            string[] fitoFharmRx = { "ua_rx", "center_rx", "south_rx", "east_rx", "south_east_rx", "west_rx", "nord_east_rx" };

            foreach (var user in fitoFharmRx)
            {
                fitoFharmMethods.LoginPlanirovschik(fitofharm, user, "1");
                Console.WriteLine("User - " + user + ":");
                Waiting.WaitForAjax();
                fitoFharmMethods.ChainsAccept();   
                fitoFharmMethods.LogoutFitoFharm(lpgoutFitoFarm);
            }
            firefox.Quit();
        }
    }
}
