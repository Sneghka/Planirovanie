using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace Planirovanie
{
    [TestFixture]
    public class CheckGeneralPlan
    {

        const string test = "http://test.stada.bi.morion.ua/";
        const string logoutTest = "http://test.stada.bi.morion.ua/logout.php";
        const string dev = "http://dev.stada.bi.morion.ua/";
        const string logoutDev = "http://dev.stada.bi.morion.ua/logout.php";
        const string stada = "http://stada.bi.morion.ua";
        const string logoutStada = "http://stada.bi.morion.ua/logout.php";


        [Test]
        public void CheckPreparationsName()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 10, 11, 12 };

            method.StoreExcelData();
            Debug.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.StorePreparationNamesFromPlanirovschik();
            Debug.WriteLine("Planirovschik was stored");
            Debug.WriteLine("Сравниваем ексель с планировщиком");
            method.CompareExcelWithWeb(months);
            Debug.WriteLine("Сравниваем Планировщик с екселем");
            method.CompareWebWithExcel(months);

            firefox.Quit();
        }

        [Test]
        public void CheckPreparationsData()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 7, 8, 9, 10, 11, 12 };

            method.StoreExcelData();
            Debug.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.CheckPreparationData(months);
            firefox.Quit();
        }

        [Test]
        public void CheckPreparationsListForProductManager()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);

            int[] chainPM = { 2200, 1965, 2718, 8030, 1901, 2195, 1590, 1763, 2128, 2494, 8003, 1172, 2708, 1638, 1174, 2393, 1788, 2113, 2222, 2711, 2149, 2205, 1514, 1598, 754, 8061, 8012, 8013};

            method.StoreExcelData();
            foreach (var user in chainPM)
            {
                method.LoginStada(test, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                method.WaitForAjax();
                method.CheckPreparationListForPM(user);
                method.LogoutStada(logoutTest);
            }
            firefox.Quit();
        }

        [Test]
        public void CheckTwoExcelFilesTerrotorii()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            method.StoreExcelDataTerritoriiPlanirovschik();
            method.StoreExcelDataTerritoriiSpravochnik();
            method.CompareTerritoriiSpravochnikWithPlanirovschik();
            method.CompareTerritoriiPlanirovschikWithSpravochnik();
            method.WriteDataToExcel();
            firefox.Quit();


        }

        [Test]
        public void CheckTwoExcelFilesBuId()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            method.StoreExcelDataBuIdPlanirovschik();
            method.StoreExcelDataBuIdSpravochnik();
            method.CompareBuIdPlanirovschikWithSpravochnik();
            method.CompareBuIdSpravochnikWithPlanirovschik();
            firefox.Quit();
        }

        [Test]
        public void CheckDistributionWithExcel()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);

            // 1340, 2711, 2149, 2205, 1514, 20
            method.StoreExcelDataDistribution();
            method.LoginStada(test, "user_20", "1");
            method.CheckDistributionDataWithExcel();

            firefox.Quit();


        }

        [Test]
        public void CheckAuditWithExcel()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);

            method.StoreExcelDataAudit();
            method.LoginStada(test, "user_1340", "1");
            method.CheckAuditDataWithExcel();
            firefox.Quit();
        }

        [Test]
        public void CheckAuditWithWeb()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            WebDriverWait wait = new WebDriverWait(firefox, TimeSpan.FromSeconds(120));

            method.LoginDashBoardAudit();
            method.SetUpFilterDashBoardAudit();
            method.LoginStadaAudit(test, "user_1340", "1");
            method.CheckAuditDataWithDashBoard();

            firefox.Quit();
        }

        [Test]
        public void CheckLoginPasswordPlanirovschik()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            var pageElements = new PageElements(firefox);
            WebDriverWait wait = new WebDriverWait(firefox, TimeSpan.FromSeconds(120));

            method.StoreLoginPasswordFromExcel();
            //method.CheckLoginPasswordMethod1(stada);
            method.CheckLoginPasswordMethod2(stada, logoutStada);

            firefox.Quit();
        }

        [Test]
        public void ClickChainsAccept()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);



            string[] chain1340 = new string[] { "1340" };
            string[] chain88 = new string[] { "2200", "1965", "2718", "625", "116", "968", "589", "419", "245", "1097", "2575", "9034", "9010" };
            string[] chain102 = new string[] { "2369", "2470", "2716", "236", "2534", "2762", "233", "8023", "8007", "8008", "8009", "9008", "9044" };
            string[] chain31_94 = new string[] { "1901", "2195", "1590", "1763", "2128", "2494", "1048", "578", "832", "874", "833", "2359", "271", "892", "8002", "9036", "9014", "9054" };
            string[] chain32 = new string[] { "8003", "718", "122", "772", "335", "8083", "83", "359", "115", "9012", "9037" };
            string[] chain42_106 = new string[] { "1172", "2708", "1638", "1174", "2111", "2393", "12", "1453", "8", "269", "64", "19", "125", "344", "9002", "9027" };
            string[] chain76 = new string[] { "1788", "2113", "2222", "2755", "253", "1801", "1798", "8025", "1835", "1525", "299", "9016", "9042" };
            string[] chain105 = new string[] { "2711", "2149", "2205", "1514", "20", "829", "220", "1235", "8005", "46", "623", "828", "36", "9019" };
            string[] chain114 = new string[] { "1598", "2212", "2725", "8022", "200", "1224", "1224", "1226", "1821", "1874", "951", "9006", "9040" };
            string[] chain112 = new string[] { "754", "8061", "1404", "8018", "647", "8017", "8019", "579", "8016", "1855", "9045", "9046" };
            string[] chain33 = new string[] { "8012", "93", "147", "1139", "1415", "1419", "142", "8015", "9048", "9023" };
            string[] chain67 = new string[] { "8013", "8070", "747", "2433", "8040", "8038", "1470", "8037", "8029", "9050", "9025" };
            string[] chain1111 = new string[] { "8071", "8072", "8073", "8074", "8075", "8076", "8077", "8078", "8079", "9051", "9052" };

            foreach (var user in chain102)
            {
                method.LoginStada(test, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                method.WaitForAjax();
                method.ChainsAccept();   // Добавить в метод - ПРОВЕРКУ ИЗМЕНЕННИЯ КНОПКИ с "Утвердить" на "Утверждён"
                method.LogoutStada(logoutTest);
            }

            firefox.Quit();
        }

        [Test]
        public void ClickChainsApprove()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            var pageElements = new PageElements(firefox);
            WebDriverWait wait = new WebDriverWait(firefox, TimeSpan.FromSeconds(120));


            string[] chain1340 = new string[] { "1340" };
            string[] chain88 = new string[] { "2200", "1965", "2718", "625", "116", "968", "589", "419", "245", "1097", "2575", "9034", "9010" };
            string[] chain102 = new string[] { "2369", "2470", "2716", "236", "2534", "2762", "233", "8023", "8007", "8008", "8009", "9008", "9044" };
            string[] chain31_94 = new string[] { "1901", "2195", "1590", "1763", "2128", "2494", "1048", "578", "832", "874", "833", "2359", "271", "892", "8002", "9036", "9014", "9054" };
            string[] chain32 = new string[] { "8003", "718", "122", "772", "335", "8083", "83", "359", "115", "9012", "9037" };
            string[] chain42_106 = new string[] { "1172", "2708", "1638", "1174", "2111", "2393", "12", "1453", "8", "269", "64", "19", "125", "344", "9002", "9027" };
            string[] chain76 = new string[] { "1788", "2113", "2222", "2755", "253", "1801", "1798", "8025", "1835", "1525", "299", "9016", "9042" };
            string[] chain105 = new string[] { "2711", "2149", "2205", "1514", "20", "829", "220", "1235", "8005", "46", "623", "828", "36", "9019" };
            string[] chain114 = new string[] { "1598", "2212", "2725", "8022", "200", "1224", "1224", "1226", "1821", "1874", "951", "9006", "9040" };
            string[] chain112 = new string[] { "754", "8061", "1404", "8018", "647", "8017", "8019", "579", "8016", "1855", "9045", "9046" };
            string[] chain33 = new string[] { "8012", "93", "147", "1139", "1415", "1419", "142", "8015", "9048", "9023" };
            string[] chain67 = new string[] { "8013", "8070", "747", "2433", "8040", "8038", "1470", "8037", "8029", "9050", "9025" };
            string[] chain1111 = new string[] { "8071", "8072", "8073", "8074", "8075", "8076", "8077", "8078", "8079", "9051", "9052" };

            foreach (var user in chain102)
            {
                method.LoginStada(test, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                method.WaitForAjax();
                method.ChainsApprove();   // Добавить в метод - ПРОВЕРКУ ИЗМЕНЕННИЯ КНОПКИ с "Утвердить" на "Утверждён"
                method.LogoutStada(logoutTest);
            }

            firefox.Quit();
        }

        [Test]
        public void StoreGr()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            var pageElements = new PageElements(firefox);
            WebDriverWait wait = new WebDriverWait(firefox, TimeSpan.FromSeconds(120));

            string[] chainPM_BU1340 = new string[] { "1340" };
            string[] chainPM_BU88 = new string[] { "2200", "1965", "2718" };
            string[] chainPM_BU102 = new string[] { "8030" };
            string[] chainPM_BU84 = new string[] { "1901", "2195", "1590", "1763", "2128", "2494" };
            string[] chainPM_BU32 = new string[] { "8003" };
            string[] chainPM_BU43 = new string[] { "1172", "2708", "1638", "1174", "2393" };
            string[] chainPM_BU76 = new string[] { "1788", "2113", "2222" };
            string[] chainPM_BU105 = new string[] { "2711", "2149", "2205", "1514" };
            string[] chainPM_BU115 = new string[] { "1598" };
            string[] chainPM_BU112 = new string[] { "754", "8061" };



            foreach (var user in chainPM_BU1340)
            {
                method.LoginStada(test, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                method.WaitForAjax();
                method.StoreGr();
                method.PrintGR();
                method.LogoutStada(logoutTest);
            }

            firefox.Quit();
        }
    }
}

