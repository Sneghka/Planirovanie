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
    public class CheckPlanirovschikStada
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
            int[] months = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };

            method.StoreExcelData(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник21.11.16.xlsx");
            Debug.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.StorePreparationNamesFromPlanirovschik();
            Debug.WriteLine("Planirovschik was stored");
            Debug.WriteLine("Данные есть в справочинке, но отсутсвуют в планировщике планировщиком");
            method.CompareExcelWithWeb(months); //STADA
            Debug.WriteLine("Данные есть в планировщике, но отсутствуют в справочнике");
            method.CompareWebWithExcel(months);//STADA

            firefox.Quit();
        }
        [Test]
        public void CheckPreparationsNameWithAdditionalInformation()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };

            method.StoreExcelData(@"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник2_05.12.16___custom_.xls");
            Console.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.StorePreparationNamesFromPlanirovschik();
            Console.WriteLine("Planirovschik was stored");
            method.ComparePreparationNameThroughObjects(months);

            firefox.Quit();
        }

        [Test]
        public void CheckPreparationsForAutoPlan()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };

            method.StoreExcelData(@"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник2_05.12.16___custom_.xls");
            Console.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.StorePreparationNamesFromPlanirovschik();
            Console.WriteLine("Planirovschik was stored");
            method.ComparePreparationWithAutoPlan(months);

            firefox.Quit();

        }

        [Test]
        public void CheckPreparationsData()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };

            method.StoreExcelData(@"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник3_14.12.16___custom_.xls");
            Debug.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.CheckPreparationData(months);
            firefox.Quit();
        }

        [Test]
        public void CheckPreparationsDataByQRT()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 1, 2, 3 };

            method.StoreExcelData(@"D:\Sneghka\IT\Selenium\Planirovschik_file\1_для_модуля__Справочник3_14.12.16___custom_.xls");
            Debug.WriteLine("Excel was stored");
            method.LoginStada(test, "user_1340", "1");
            method.CheckPreparationDataByQrt(months);
            firefox.Quit();
        }

        [Test]
        public void CheckPreparationsDataByUsers() //проверка и по году и по кварталу
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] months = { 1, 2, 3 };
            method.StoreExcelData(@"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник3_14.12.16___custom_.xls");
            Debug.WriteLine("Excel was stored");
            /* var users = new int[] {1340};*/
            var users = new int[] { 2200, 1965, 2718, 2864, 1901, 2195, 2128, 2494, 8003, 2859, 2708, 1638, 1174, 2113, 2711, 8067, 2205, 1598, 2212, 754, 2849, 2861, 8061, 8012, 8013,  8071};//2494


            foreach (var user in users)
            {
                method.LoginStada(test, "user_"+user, "1");
                Console.WriteLine("User_" + user + ":");
                method.CheckPreparationDataByUserGlobal(months,user);
                method.LogoutStada(logoutTest);
            }
            firefox.Quit();
        }

        [Test]
        public void CheckPreparationsListForProductManager()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);


            /* int[] chainPM = { 58, 754, 1167, 1174, 1598, 1638, 1901, 1965, 2071, 2113, 2128, 2195, 2200, 2205, 2206, 2212, 2494, 2708, 2711, 2718, 2849, 2859 };*/

            int[] chainPM = { 8003, 754, 8061, 1174, 1598, 1638, 1901, 1965, 8012, 8013, 2113, 2128, 2195, 2200, 2205, 2212, 2494, 2708, 2711, 2718, 2849, 2859, 2861, 2864,  8067 };

            method.StoreExcelData(@"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник2_05.12.16___custom_.xls");
            foreach (var user in chainPM)
            {

                method.LoginStada(test, "user_" + user, "1");
                if (!method.IsLoginSuccess(test, user.ToString(), "1"))
                {
                    Console.WriteLine("user_" + user + "  Incorrect login or password");
                    firefox.Navigate().GoToUrl(logoutTest);
                    continue;
                }
                if (!method.IsPreparationListExist())
                {
                    firefox.Navigate().GoToUrl(logoutTest);
                    continue;
                }

                Console.WriteLine("User_" + user + ":");
                Waiting.WaitForAjax(firefox);
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

            // 1340, 9010_m004c
            //МОСКВА ----- 9010_m004c, 9012_l428n, 9002_b155d, 9016_v827s, 9045_y374c, 9025_l043g
            method.StoreExcelDataDistribution();
            method.LoginStada(stada, "user_9045", "y374c");
            method.CheckDistributionDataWithExcel();

            firefox.Quit();
        }

        [Test]
        public void CheckDistributionWithExcelGlobalRussia()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            method.StoreExcelDataDistributionFromSpravochnik(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник2_05.12.2016.xlsx", "New");
            method.StoreExcelDataDistributionFromSpravochnik(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник3_14.16.16.xlsx", "Факт янв-окт 2016");
            method.LoginStada(stada, "user_1340", "m600e");
            method.CheckDistributionDataWithExcelRussia();

        }

        [Test]
        public void CheckDistributionWithExcelForNop()
        {
            int[] chainNOP = {625, 1048, 718, 12, 2755, 20, 2725,1404, 8069, 8070}; 
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            
            method.StoreLoginPasswordFromExcel();// @"D:\Sneghka\Selenium\Projects\Planirovschik\Check_Login_Pass.xlsx"
            method.StoreExcelDataDistributionFromSpravochnik(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник2_05.12.2016.xlsx", "New");
            method.StoreExcelDataDistributionFromSpravochnik(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник3_14.16.16.xlsx", "Факт янв-окт 2016");
            foreach (var nop in chainNOP)
            {
                var user = "user_" + nop;
                var password = method.GetPasswordByUser(nop.ToString());
                Console.WriteLine(user + "_" + password);
                method.LoginStada(stada, user, password); 
                method.CheckDistributionDataWithExcelForNop();
                method.LogoutStada(logoutStada);
            }
        }

        [Test]
        public void CheckDistributionWithExcelForTM()
        {
            //Regions - Центр, Юг, Урал, Поволжье, Москва, Северо-Запад, Сибирь-Дальний Восток

            int[] TmCenter = { 116, 892, 115, 551, 1525, 1235, 1874, 8018, 93, 747, 8073 }; //Центр
            int[] TmNorthWest = {968,832,2351, 8083, 64, 1835, 220, 1224, 8019, 1415, 8038, 8076};//Северо-Запад
            int[] TmUral = {589,833,359,344,8039,36,8022,647,147,2433,8072 };//Урал
            int[] TmPovolzhe = { 419, 578, 772, 2534, 299, 46, 623, 1226, 8016, 142, 8037};//Поволжье
            int[] TmMoscow = { 245, 874,335,8002, 2762,8025,8005,200,1855,8015,8029,8079 };//Москва
            int[] TmSibirDalnVostok = { 1097, 2359, 83, 269, 253, 828, 1821, 8017, 1139 ,8070, 8072 };//Сибирь-Дальний Восток
            int[] TmSouth = { 937, 829, 271, 19, 1801, 122, 579,1419,951,1470,8077 };//Юг


            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);

            method.StoreLoginPasswordFromExcel();
            method.StoreExcelDataDistributionFromSpravochnik(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник2_05.12.2016.xlsx", "New");
            method.StoreExcelDataDistributionFromSpravochnik(@"D:\Sneghka\Selenium\Projects\Planirovschik\Справочник3_14.16.16.xlsx", "Факт янв-окт 2016");
            foreach (var nop in TmCenter)
            {
                var user = "user_" + nop;
                var password = method.GetPasswordByUser(nop.ToString());
                Console.WriteLine(user + "_" + password);
                method.LoginStada(stada, user, password);
                method.CheckDistributionDataWithExcelForTm("Центр");
                method.LogoutStada(logoutStada);
            }
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
        public void CheckAuditWithExcelForPm()
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            int[] chainPM = { 2200, 1965, 2718, 2864, 1901, 2195, 2128, 2494, 8003, 2859, 2708, 1638, 1174, 2113, 2711, 8067,  2205, 1598, 2212, 754, 2849, 2861, 8061, 8012, 8013};
            method.StoreExcelDataAudit();


            foreach (var user in chainPM)
            {

                method.LoginStada(test, "user_" + user, "1");
                if (!method.IsLoginSuccess(test, user.ToString(), "1"))
                {
                    Console.WriteLine("user_" + user + "  Incorrect login or password");
                    firefox.Navigate().GoToUrl(logoutTest);
                    continue;
                }
                if (!method.IsPreparationListExist())
                {
                   firefox.Navigate().GoToUrl(logoutTest);
                    continue;
                }

                Console.WriteLine("User_" + user + ":");
                Waiting.WaitForAjax(firefox);
                method.CheckAuditDataWithExcelForPm();
                method.LogoutStada(logoutTest);
            }
           
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

            //new PM    8003, 754, 8061, 1174, 1598, 1638, 1901, 1965, 8012, 8013, 2113, 2128, 2195, 2200, 2205, 2212, 2494, 2708, 2711, 2718, 2849, 2859, 2861, 2864,  8067 

            string[] chain1340 = new string[] { "1340" };
            string[] chain88 = new string[] {/* "2200", "1965", "2718",*/ "625", "116", "968", "589", "419", "245", "1097","937",  "9034", "9010" };// ГОТОВО
            string[] chain31_94 = new string[] { /*"2864","1901", "2195", "2128", "2494", "1048", "578", "832", "2351","874", "833", "2359", "271",*/ /*"892"*//*, "8002", "9036","9056",*/ "9014", "9054" };// УТВЕРДИЛА
            string[] chain32 = new string[] { /*"8003",*//* "718", "122", "772", "335", "8083", "83", "359", "115",*/ "9012", "9037" };//Готово
            string[] chain42_106 = new string[] { "2859", "2708", "1638", "1174","12", "551", "2762", "269","64", "19", "2534", "344", "9002", "9027" };//
            string[] chain76 = new string[] { /*"2113", "2755", "253", "1801", "8039", "8025", "1835",*/ "1525", "299", "9016", "9041"};//
            string[] chain105 = new string[] { /*"2711", "8067", "2205",*/ /*"20",*/ /*"829", "220", "1235", "8005", "46",*/ "623", "828", "36", "9030", "9020" };//
            string[] chain115 = new string[] { "1598", "2212", "2725", "8022", "200", "1224", "1226", "1821", "1874", "951", "9006", "9040" };//
            string[] chain116 = new string[] { "754", "2849","2861","8061", "1404", "8018", "647", "8017", "8019", "579", "8016", "1855", "9045", "9046" }; //
            string[] chain33 = new string[] { /*"8012",*/ "8069", "93", "147", "1139", "1415", "1419", "142", "8015", "9048", "9023" };// 
            string[] chain67 = new string[] { "8013", "8070", "747", "2433", "8040", "8038", "1470", "8037", "8029", "9050", "9025" }; // ГОТОВО

            string[] chain1111 = new string[] { "8071", "8072", "8073", "8074", "8075", "8076", "8077", "8078", "8079", "9051", "9052" };

            foreach (var user in chain88)
            {
                method.LoginStada(test, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                Waiting.WaitForAjax(firefox);
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

            string[] chain88 = new string[] { "9010", "9034", "937", "1097", "245", "419", "589", "968", "116"};//ГОТОВО
            string[] chain31_94 = new string[] { "9054", "9014", "9056", "9036", "892", "271", "2359", "833", "874", "2351", "832", "8002", "578" };//УТВЕРДИЛА
            string[] chain32 = new string[] { "9037", "9012", "115", "359", "83", "8083", "335", "772", "122"};//УТВЕРДИЛА
            string[] chain42_106 = new string[] { "9027", "9002", "344", "2534", "19", "64", "269", "2762", "551" };//автотест
            string[] chain76 = new string[] { /*"9041", "9016",*/ "299", "1525", "1835", "8025", "8039", "1801", "253"};
            string[] chain105 = new string[] { "9020", "9030", "36", "828", "623", "46", "8005", "1235", "220", "829"};
            string[] chain115 = new string[] { "9040", "9006", "951", "1874", "1821", "1226", "1224", "200", "8022"};
            string[] chain116 = new string[] { "9046", "9045", "1855", "8016", "579", "8019", "8017", "647", "8018"};
            string[] chain33 = new string[] { /*"9023",*/ "9048", "8015", "142", "1419", "1415", "1139", "147", "93" };//
            string[] chain67 = new string[] { "9025", "9050", "8029", "8037", "1470", "8038", "8040", "2433", "747" };//ГОТОВО

            string[] chain43 = new string[] { "9002", "9027", "344", "19", "64", "269", "1453", "8","125", }; // бывший 42_106 (в прошлый период - осень)
            
           

            foreach (var user in chain88)
            {
                method.LoginStada(test, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                Waiting.WaitForAjax(firefox);
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
                Waiting.WaitForAjax(firefox);
                method.StoreGr();
                method.PrintGR();
                method.LogoutStada(logoutTest);
            }

            firefox.Quit();
        }
       

        [Test]
        public void IsGrChangesPossibleForUser() // проверяет выключены ли некоторые препараты , для НОП которым  всё разрешено
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            
            string[] chainPM_BU84_32_43 = new string[] { "1901", "2195", "1590", "1763", "2128", "2494", "8003", "1172", "2708", "1638", "1174", "2393" };
            string[] chainPM_BU88 = { "2200", "1965", "2718" };
            string[] chainPM_BU76 = { "1788", "2113", "2222" };
            string[] chainPM_BU105_112_115 = { "2711", "2149", "2205", "1514", "754", "8061", "1598" };
            string[] chainPM_NOP = { "625", "1048", "718", "12", "2755", "1404", "20", "1404", "2725", "8069", "8070" };

            method.StoreExcelDataAny(@"D:\Sneghka\Selenium\Projects\Planirovschik\Замороженные_препараты_на_НОП.xls");
            /*method.StoreLoginPasswordFromExcel();*/

            foreach (var user in chainPM_NOP)
            {

                /* method.LoginStada(stada, "user_" + user, method.GetPasswordByUser(user));*/
                method.LoginStada(stada, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                Waiting.WaitForAjax(firefox);
                method.IsGrUnchangeable();

                method.LogoutStada(logoutStada);
            }

            firefox.Quit();
        }

        [Test]
        public void IsGRChangesImpossibleForUser() // проверяет включены ли некоторые препараты для НОПов которым всё запрещено
        {
            var firefox = new FirefoxDriver();
            var method = new Methods(firefox);
            var pageElements = new PageElements(firefox);
            WebDriverWait wait = new WebDriverWait(firefox, TimeSpan.FromSeconds(120));

            string[] chainPM_BU84_32_43 = new string[] { "1901", "2195", "1590", "1763", "2128", "2494", "8003", "1172", "2708", "1638", "1174", "2393" };
            string[] chainPM_BU88 = { "2200", "1965", "2718" };
            string[] chainPM_BU76 = { "1788", "2113", "2222" };
            string[] chainPM_BU105_112_115 = { "2711", "2149", "2205", "1514", "754", "8061", "1598" };
            string[] chainPM_NOP = { "625", "1048", "718", "12", "2755", "1404", "20", "2725", "8069", "8070" };

            method.StoreExcelDataAny(@"D:\Sneghka\Selenium\Projects\Planirovschik\Unfrozen.xls");

            foreach (var user in chainPM_NOP)
            {
                method.LoginStada(stada, "user_" + user, "1");
                Console.WriteLine("User_" + user + ":");
                Waiting.WaitForAjax(firefox);
                method.IsGrChangeable();
                method.LogoutStada(logoutStada);
            }
            firefox.Quit();
        }
    }
}

