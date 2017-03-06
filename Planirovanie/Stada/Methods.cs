using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.IO.Pipes;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Planirovanie.CheckStadaPlan;
using Planirovanie.Objects;
using HtmlAgilityPack;

namespace Planirovanie
{
    public static class StringExtensions
    {
        public static bool ContainsIgnoreCase(this string source, string toCheck)
        {
            return source.IndexOf(toCheck, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }
    }

    public class Methods
    {
        private readonly FirefoxDriver _firefox;

        private int _numberTableRows;
        private RowDataList _preparationNamePlanirovschik = new RowDataList();
        private RowDataList _preparationDataSpravochnik = new RowDataList();
        private RowTerritoriiList _planirovschikTerritorii = new RowTerritoriiList();
        private RowTerritoriiList _spravochnikTerritorii = new RowTerritoriiList();
        private RowTerritoriiList planirovschikBuId = new RowTerritoriiList();
        private RowTerritoriiList spravochnikBuId = new RowTerritoriiList();
        private List<RowTerritorii> _differencePlanirovschikWithSpravochik = new RowTerritoriiList();
        private List<RowTerritorii> _differenceSpravochikWithPlanirovschik = new RowTerritoriiList();
        private RowDataList _distribution2016XlsList = new RowDataList();
        private RowDataList _audit2015XlsList = new RowDataList();
        //private List<string> _handles;
        private string _planirovschikdWindow;
        private string _dashBoardWindow;
        private List<LoginPassword> _loginPasswordList = new List<LoginPassword>();
        private List<string> _grListValue = new List<string>();
        private DistributionSpravochnikRowList _distributionSpravochnikRows = new DistributionSpravochnikRowList();
        private List<User> _usersListPlanirovschik = new List<User>();
        private UserList _usersListForEmailSpravochnik = new UserList();
        private List<PlanTableRow> _planForLgotaBu33 = new List<PlanTableRow>();
        private readonly string regionsString = "Юг,Центр,Поволжье,Урал,Москва,Северо-Запад,Сибирь-Дальний Восток";



        public FirefoxProfile SetUpFirefoxProfile()
        {
            var downloadPath = @"D:\DownloadTest";
            FirefoxProfile firefoxProfile = new FirefoxProfile();
            firefoxProfile.SetPreference("browser.download.folderList", 2);
            firefoxProfile.SetPreference("browser.download.dir", downloadPath);
            // firefoxProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
            firefoxProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            return firefoxProfile;
        }

        public Methods(FirefoxDriver firefox)
        {
            _firefox = firefox;
        }

        #region Compare Name and Data of Preparations

        public void StoreExcelData(string path) //@"D:\Sneghka\Selenium\Projects\Planirovschik\GP_24.08.2016.xlsx"
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path,
                "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var name = row["Name"].ToString().Trim().Replace("\u00A0", " ").ToLower();
                var rowData = new RowData
                {
                    IdPrUniq = Convert.ToInt32(row["id_PrUniq"]),
                    Name = Regex.Replace(name, @"\s+", " "),
                    Id_BU = Convert.ToInt32(row["id_BU"]),
                    BusinessUnit = row["BU"].ToString(),
                    Year = Convert.ToInt32(row["Год"]),
                    Month = Convert.ToInt32(row["Месяц"]),
                    Segment = Convert.ToInt32(row["Сегмент"]),
                    Upakovki = Convert.ToInt32(row["Сумма в упаковках"]),
                    Summa = Convert.ToDecimal(row["Сумма в рублях"]),
                    Group = row["Группа"].ToString(),
                    IdSotr = Convert.ToInt32(row["id_Sotr"])
                };
                _preparationDataSpravochnik.Add(rowData);
            }
        }

        public void StoreExcelDataBuTerritorii(string path)
        //@"D:\Sneghka\Selenium\Projects\Planirovschik\FitoPharm.xlsx"
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path,
                "Select * from [Зоны ответственности$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var rowData = new RowTerritorii()
                {
                    IdSotr = Convert.ToInt32(row["id_Sotr"]),
                    FIO = row["ФИО"].ToString(),
                    Position = row["Должность"].ToString(),
                    BuId = Convert.ToInt32(row["BUID"]),
                    //Name1RegionId = row["id региона"].ToString(),
                    Name1RegionName = row["Регион"].ToString(),



                };
                _spravochnikTerritorii.Add(rowData);
            }
        }

        public void StoreExcelDataAny(string path) //@"D:\Sneghka\Selenium\Projects\Planirovschik\FitoPharm.xlsx"
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path,
                "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                /*var name = row["Полное наименование"].ToString().Trim().ToLower();*/
                var name = row["Name"].ToString().Trim().Replace("\u00A0", " ").ToLower();
                var rowData = new RowData
                {

                    Name = Regex.Replace(name, @"\s+", " "),
                    Id_BU = Convert.ToInt32(row["id_BU"]),
                    IdPrUniq = Convert.ToInt32(row["Id_PrUniq"]),
                    /*Segment = Convert.ToInt32(row["Сегмент"])*/
                };
                _preparationDataSpravochnik.Add(rowData);
            }

        }

        public void LoginStada(string url, string login, string password)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            _firefox.Navigate().GoToUrl(url);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
            pageElements.LoginField.SendKeys(login);
            pageElements.PasswordField.SendKeys(password);
            pageElements.SubmitButton.Click();
            Thread.Sleep(2000);
        }

        public bool IsLoginSuccess(string url, string login, string password)
        {
            if (_firefox.FindElement(By.XPath(".//*[@id='dialog-confirm']")).GetAttribute("style") == "display: none;")
                return true;
            return false;
        }

        public bool IsPreparationListExist()
        {
            if (Helper.IsElementPresent(By.XPath("html/body/div[4]/div[3]/div/button[1]"), _firefox))
                //кнопка "Закрыть" на списке препаратов
                return true;
            return false;
        }

        public void StorePreparationNamesFromPlanirovschik()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(1000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            _numberTableRows = tableRows.Count;
            Thread.Sleep(4000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице Планировщика");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                var name =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]"))
                        .Text.Trim()
                        .Replace("\u00A0", " ")
                        .ToLower();

                var rowData = new RowData()
                {
                    IdPrUniq = Convert.ToInt32(tableRows[i - 1].GetAttribute("data_id")), // add preparation id
                    Id_BU = Convert.ToInt32(tableRows[i - 1].GetAttribute("bu_id")), // add preparation bu_id
                    Name = Regex.Replace(name, @"\s+", " "), // add preparation name
                    Status =
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[3]"))
                            .GetAttribute("aria-disabled")

                };
                _preparationNamePlanirovschik.Add(rowData);
            }
        }

        public void GetListPreparationFromExcel(int[] months)
        {
            Debug.WriteLine("Список препаратов из Екселя:");
            foreach (var name in _preparationDataSpravochnik.GetUniqueWebNames(months))
            {
                Debug.WriteLine(name);
            }
        }

        public void ComparePreparationNameThroughObjects(int[] months)
        {
            var convertSpravochnik = RowDataList.ConvertSpravochikList(months, _preparationDataSpravochnik);
            Console.WriteLine("СПРАВОЧНИК");

            var diff1 = RowDataList.CompareRowDataObjects(convertSpravochnik, _preparationNamePlanirovschik);
            if (diff1.Count != 0)
            {
                Console.WriteLine("Данные из справочника отсутствуют в планировщике:");
                foreach (var d in diff1)
                {
                    Console.WriteLine(d.IdPrUniq + " " + d.Name + " (BU_ID - " + d.Id_BU + "; Segment - " + d.Segment +
                                      "; Group - " + d.Group + ")");
                }
            }
            else
            {
                Console.WriteLine("Сверка справочника с планировщиком. Расхождений нет");
            }
            var diff2 = RowDataList.CompareRowDataObjects(_preparationNamePlanirovschik, convertSpravochnik);
            if (diff2.Count != 0)
            {
                Console.WriteLine("Данные из планировщика отсутствуют в справочнике:");
                foreach (var d in diff2)
                {
                    Console.WriteLine(d.IdPrUniq + " " + d.Name + " (BU_ID - " + d.Id_BU + ")");
                }
            }
            else
            {
                Console.WriteLine("Сверка планировщика со справочником. Расхождений нет");
            }

        }

        public void ComparePreparationWithAutoPlan(int[] months)
        {
            var convertSpravochnik = RowDataList.ConvertSpravochikList(months, _preparationDataSpravochnik);
            var convertSpravochnikWithAutoplanOnly =
                RowDataList.GetPreparationWithAutoPlanFromSpravochnik(convertSpravochnik);
            var planirovschikWithAutoplanOlny =
                RowDataList.GetPreparationWithAutoPlanFromPlanirovschik(_preparationNamePlanirovschik);


            var diff1 = RowDataList.CompareRowDataObjects(convertSpravochnikWithAutoplanOnly,
                planirovschikWithAutoplanOlny);
            if (diff1.Count != 0)
            {
                Console.WriteLine("Данные из справочника отсутствуют в планировщике:");
                foreach (var d in diff1)
                {
                    Console.WriteLine(d.IdPrUniq + " " + d.Name + " (BU_ID - " + d.Id_BU + "; Segment - " + d.Segment +
                                      "; Group - " + d.Group + ")");
                }
            }
            else
            {
                Console.WriteLine("Сверка справочника с планировщиком. Расхождений нет");
            }
            var diff2 = RowDataList.CompareRowDataObjects(planirovschikWithAutoplanOlny,
                convertSpravochnikWithAutoplanOnly);
            if (diff2.Count != 0)
            {
                Console.WriteLine("Данные из планировщика отсутствуют в справочнике:");
                foreach (var d in diff2)
                {
                    Console.WriteLine(d.IdPrUniq + " " + d.Name + " (BU_ID - " + d.Id_BU + ")");
                }
            }
            else
            {
                Console.WriteLine("Сверка планировщика со справочником. Расхождений нет");
            }
        }

        public void CompareWebWithExcel(int[] months)
        {
            var difference = RowDataList.CompareStrings(_preparationNamePlanirovschik.GetNamesList(),
                _preparationDataSpravochnik.GetUniqueWebNames(months));
            if (difference.Count == 0)
            {
                Debug.WriteLine("Расхождений нет");
            }
            else
            {
                foreach (var x in difference)
                {
                    Debug.WriteLine(x);
                }
            }
        }

        public void CompareExcelWithWeb(int[] months)
        {
            var difference = RowDataList.CompareStrings(_preparationDataSpravochnik.GetUniqueWebNames(months),
                _preparationNamePlanirovschik.GetNamesList());


            Debug.WriteLine(difference.Count + "Count");
            if (difference.Count == 0)
            {
                Debug.WriteLine("Расхождений нет");
            }
            else
            {
                foreach (var x in difference)
                {
                    Debug.WriteLine(x);

                }
            }
        }

        public void MessageCheckPreparationMethodByMonth(int month, int pcsSpravochnik, int pcsPlanirovschik)
        {
            if (pcsSpravochnik == pcsPlanirovschik)
            {
                Console.WriteLine(month + "month pcs: справочник" + pcsSpravochnik + " = " + pcsPlanirovschik +
                                  " планировщик");
            }
            else
            {
                Console.WriteLine(month + "month pcs: справочник" + pcsSpravochnik + " НЕ РАВНО!!! " + pcsPlanirovschik +
                                  " планировщик");
            }
        }

        public void MessageCheckPreparationMethodTotal(int preparationId, int[] months, decimal totalSum, int totalPcs)
        {
            if ((_preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) - totalSum) < 5 &&
                _preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) - totalSum > -5)
            {
                Console.WriteLine("Total sum: справочник " +
                                  _preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) + " = " +
                                  totalSum + " планировщик");
            }
            else
            {
                Console.WriteLine("Total sum: справочник " +
                                  _preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) +
                                  " НЕ РАВНО !!!! " + totalSum + " планировщик");
            }

            if (_preparationDataSpravochnik.GetTotalPcsById(preparationId, months) == totalPcs)
            {
                Console.WriteLine("Total pcs: справочник " +
                                  _preparationDataSpravochnik.GetTotalPcsById(preparationId, months) + " = " + totalPcs +
                                  " планировщик");
            }
            else
            {
                Console.WriteLine("Total pcs: справочник " +
                                  _preparationDataSpravochnik.GetTotalPcsById(preparationId, months) + " НЕ РАВНО!!! " +
                                  totalPcs + " планировщик");
            }
        }

        public void CheckPreparationData(int[] months)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= _numberTableRows; i++)
            {
                Console.WriteLine("№" + i);
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationBuId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("bu_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
                Thread.Sleep(1000);

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("Кнопка расчёт неактивна - " + preparationName);
                    continue;
                }

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                Thread.Sleep(2000);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TotalPcsXPath)));

                if (preparationId < 0) // проверяем является ли препарат льготным
                {
                    Console.WriteLine(preparationId + " " + preparationName + " (BU" + preparationBuId + "): ");
                    preparationId *= -1;
                    Dictionary<int, int> monthSumLgota = new Dictionary<int, int>();

                    foreach (var month in months)
                    {
                        var totalPcsMonthLgota =
                            Convert.ToInt32(pageElements.GetPcsLgotaMonth(month).Text.Replace(" ", ""));
                        monthSumLgota.Add(month, totalPcsMonthLgota);
                    }

                    foreach (KeyValuePair<int, int> kvp in monthSumLgota)
                    {
                        MessageCheckPreparationMethodByMonth(kvp.Key,
                            _preparationDataSpravochnik.GetPcsByIdAndSegmentAndMonth(preparationId, kvp.Key, 2),
                            kvp.Value);
                    }

                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                    Thread.Sleep(500);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.FindPreparationInputFieldXPath)));
                    continue;
                }
                if (preparationId > 0)
                {
                    Console.WriteLine(preparationId + " " + preparationName + " (BU" + preparationBuId + "): ");
                    decimal totalSum;
                    int totalPcs;
                    if (pageElements.TotalSumRub.Text == "0")
                    {
                        totalSum = 0;
                        totalPcs = Convert.ToInt32(pageElements.TotalPcs.Text.Replace(" ", ""));
                    }
                    else
                    {
                        totalSum =
                            Convert.ToDecimal(
                                pageElements.TotalSumRub.Text.Substring(0, pageElements.TotalSumRub.Text.Length - 5)
                                    .Replace(" ", "")
                                    .Replace(".", ","));
                        totalPcs = Convert.ToInt32(pageElements.TotalPcs.Text.Replace(" ", ""));
                    }

                    Dictionary<int, int> monthSum = new Dictionary<int, int>();

                    foreach (var month in months)
                    {
                        var totalPcsMonth = Convert.ToInt32(pageElements.GetPcsMonth(month).Text.Replace(" ", ""));
                        monthSum.Add(month, totalPcsMonth);
                    }

                    MessageCheckPreparationMethodTotal(preparationId, months, totalSum, totalPcs);

                    foreach (KeyValuePair<int, int> kvp in monthSum)
                    {
                        MessageCheckPreparationMethodByMonth(kvp.Key,
                            _preparationDataSpravochnik.GetSumPcsByIdAndMonth(preparationId, kvp.Key), kvp.Value);
                    }
                }

                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                Thread.Sleep(500);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.FindPreparationInputFieldXPath)));

            } //конец цикла FOR перебора всех препаратов

        } //конец метода

        public void CheckPreparationDataByQrt(int[] months)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            _numberTableRows = tableRows.Count;
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= _numberTableRows; i++)
            {
                Console.WriteLine("№" + i);
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationBuId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("bu_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
                Thread.Sleep(1000);

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("Кнопка расчёт неактивна - " + preparationName);
                    continue;
                }

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                Thread.Sleep(2000);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TotalPcsXPath)));

                if (preparationId < 0) // проверяем является ли препарат льготным
                {
                    Console.WriteLine(preparationId + " " + preparationName + " (BU" + preparationBuId + "): ");
                    preparationId *= -1;

                    var totalSumString = _firefox.FindElement(By.XPath(" .//*[@id='sumEuro2']"));

                    var totalPcsQrtLgota =
                        Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='sumPOPPVP2']")).Text.Replace(" ", ""));
                    decimal totalSumQrtLgota;
                    if (totalSumString.Text == "0")
                    {
                        totalSumQrtLgota = 0;
                    }
                    else
                    {
                        totalSumQrtLgota =
                            Convert.ToDecimal(
                                totalSumString.Text.Substring(0, totalSumString.Text.Length - 5)
                                    .Replace(" ", "")
                                    .Replace(".", ","));
                    }
                    var totalSumSpravochnik = _preparationDataSpravochnik.GetTotalSumRubById(preparationId, months);
                    var totalPcsSpravochnik = _preparationDataSpravochnik.GetTotalPcsById(preparationId, months);

                    if (totalSumQrtLgota - totalSumSpravochnik < 10 && totalSumQrtLgota - totalSumSpravochnik > -10)
                    {
                        Console.WriteLine(totalSumQrtLgota + " = " + totalSumSpravochnik + " (сумма за 1-ый квартал)");
                        Console.WriteLine("разница - " + (totalSumQrtLgota - totalSumSpravochnik));
                    }
                    else
                    {
                        Console.WriteLine(totalSumQrtLgota + " НЕ РАВНО!!!! " + totalSumSpravochnik +
                                          " (сумма за 1-ый квартал)");
                        Console.WriteLine("разница - " + (totalSumQrtLgota - totalSumSpravochnik));
                    }
                    if (totalPcsQrtLgota != totalPcsSpravochnik)
                    {
                        Console.WriteLine(totalPcsQrtLgota + " НЕ РАВНО!!!! " + totalPcsSpravochnik +
                                          " (сумма за 1-ый квартал)");
                    }
                    else
                    {
                        Console.WriteLine(totalPcsQrtLgota + " = " + totalPcsSpravochnik + " (сумма за 1-ый квартал)");
                        Console.WriteLine(" ");
                    }

                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                    Thread.Sleep(500);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.FindPreparationInputFieldXPath)));
                    continue;
                }
                if (preparationId > 0)
                {
                    Console.WriteLine(preparationId + " " + preparationName + " (BU" + preparationBuId + "): ");

                    var totalSumString = _firefox.FindElement(By.XPath(" .//*[@id='sumEuro2']"));

                    var totalPcsQrtLgota =
                        Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='sumPOPPVP2']")).Text.Replace(" ", ""));

                    decimal totalSumQrtLgota;
                    if (totalSumString.Text == "0")
                    {
                        totalSumQrtLgota = 0;
                    }
                    else
                    {
                        totalSumQrtLgota =
                            Convert.ToDecimal(
                                totalSumString.Text.Substring(0, totalSumString.Text.Length - 5)
                                    .Replace(" ", "")
                                    .Replace(".", ","));
                    }



                    var totalSumSpravochnik = _preparationDataSpravochnik.GetTotalSumRubById(preparationId, months);
                    var totalPcsSpravochnik = _preparationDataSpravochnik.GetTotalPcsById(preparationId, months);

                    if (totalSumQrtLgota - totalSumSpravochnik < 10 && totalSumQrtLgota - totalSumSpravochnik > -10)
                    {
                        Console.WriteLine(totalSumQrtLgota + " = " + totalSumSpravochnik + " (сумма за 1-ый квартал)");
                        Console.WriteLine("разница - " + (totalSumQrtLgota - totalSumSpravochnik));
                    }
                    else
                    {
                        Console.WriteLine(totalSumQrtLgota + " НЕ РАВНО!!!! " + totalSumSpravochnik +
                                          " (сумма за 1-ый квартал)");
                        Console.WriteLine("разница - " + (totalSumQrtLgota - totalSumSpravochnik));

                    }
                    if (totalPcsQrtLgota != totalPcsSpravochnik)
                    {
                        Console.WriteLine(totalPcsQrtLgota + " НЕ РАВНО!!!! " + totalPcsSpravochnik +
                                          " (кол-во за 1-ый квартал)");
                    }
                    else
                    {
                        Console.WriteLine(totalPcsQrtLgota + " = " + totalPcsSpravochnik + " (кол-во за 1-ый квартал)");
                    }

                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                    Thread.Sleep(500);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.FindPreparationInputFieldXPath)));

                } //конец цикла FOR перебора всех препаратов

            }
        } //конец метода

        public void CheckPreparationDataByUserGlobal(int[] months, int userId) //проверка год + квартал
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= _numberTableRows; i++)
            {
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationBuId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("bu_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton = _firefox.FindElement(By.XPath(raschetButtonXPath));
                Console.WriteLine("№" + i + ". " + preparationId + " " + preparationName + " (BU" + preparationBuId +
                                  "): ");

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("Кнопка расчёт неактивна .");
                    continue;
                }

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                Thread.Sleep(2000);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TotalPcsXPath)));
                Thread.Sleep(2000);

                if (preparationId < 0) // проверяем является ли препарат льготным
                {
                    preparationId *= -1;
                }
                decimal totalSumYearPlanirovshik;

                var totalPcsYearPlanirovshik = Convert.ToInt32(pageElements.TotalPcs.Text.Replace(" ", ""));
                if (pageElements.TotalSumRub.Text == "0")
                {
                    totalSumYearPlanirovshik = 0;
                }
                else
                {
                    totalSumYearPlanirovshik =
                        Convert.ToDecimal(
                            pageElements.TotalSumRub.Text.Substring(0, pageElements.TotalSumRub.Text.Length - 5)
                                .Replace(" ", "")
                                .Replace(".", ","));
                }

                var totalSumStringQrtPlanirovshik = _firefox.FindElement(By.XPath(" .//*[@id='sumEuro2']"));
                var totalPcsQrtPlanirovshik =
                    Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='sumPOPPVP2']")).Text.Replace(" ", ""));

                decimal totalSumQrtPlanirovschik;
                if (totalSumStringQrtPlanirovshik.Text == "0")
                {
                    totalSumQrtPlanirovschik = 0;
                }
                else
                {
                    totalSumQrtPlanirovschik =
                        Convert.ToDecimal(
                            totalSumStringQrtPlanirovshik.Text.Substring(0,
                                totalSumStringQrtPlanirovshik.Text.Length - 5).Replace(" ", "").Replace(".", ","));
                }

                var totalSumSpravochnikQrt =
                    _preparationDataSpravochnik.GetTotalSumRubByIdAndUserAndMonths(preparationId, months, userId);
                var totalPcsSpravochnikQrt = _preparationDataSpravochnik.GetTotalPcsByIdAndUserAndMonths(preparationId,
                    months, userId);
                var totalSumSpravochnikYear = _preparationDataSpravochnik.GetTotalSumByIdAndUser(preparationId, userId);
                var totalPcsSpravochnikYear = _preparationDataSpravochnik.GetTotalPcsByIdAndUser(preparationId, userId);

                //Сверка суммы руб. за год
                if (totalSumYearPlanirovshik - totalSumSpravochnikYear < 10 &&
                    totalSumYearPlanirovshik - totalSumSpravochnikYear > -10)
                {
                    Console.WriteLine(totalSumYearPlanirovshik + " = " + totalSumSpravochnikYear + " (сумма за год)" +
                                      " - " + "разница - " + (totalSumYearPlanirovshik - totalSumSpravochnikYear));
                }
                else
                {
                    Console.WriteLine(totalSumYearPlanirovshik + " НЕ РАВНО!!!! " + totalSumSpravochnikYear +
                                      " (сумма за год)" + " - " + "разница - " +
                                      (totalSumYearPlanirovshik - totalSumSpravochnikYear));
                }


                //Сверка суммы руб. за квартал
                if (totalSumQrtPlanirovschik - totalSumSpravochnikQrt < 10 &&
                    totalSumQrtPlanirovschik - totalSumSpravochnikQrt > -10)
                {
                    Console.WriteLine(totalSumQrtPlanirovschik + " = " + totalSumSpravochnikQrt +
                                      " (сумма за 1-ый квартал)" + " - " + "разница - " +
                                      (totalSumQrtPlanirovschik - totalSumSpravochnikQrt));
                }
                else
                {
                    Console.WriteLine(totalSumQrtPlanirovschik + " НЕ РАВНО!!!! " + totalSumSpravochnikQrt +
                                      " (сумма за 1-ый квартал)" + " - " + "разница - " +
                                      (totalSumQrtPlanirovschik - totalSumSpravochnikQrt));
                }


                //Сверка упаковок за год
                if (totalPcsYearPlanirovshik != totalPcsSpravochnikYear)
                {
                    Console.WriteLine(totalPcsYearPlanirovshik + " НЕ РАВНО!!!! " + totalPcsSpravochnikYear +
                                      " (упаковки за год)");
                }
                else
                {
                    Console.WriteLine(totalPcsYearPlanirovshik + " = " + totalPcsSpravochnikYear + " (упаковки за год)");
                }

                //Сверка упаковок за квартал
                if (totalPcsQrtPlanirovshik != totalPcsSpravochnikQrt)
                {
                    Console.WriteLine(totalPcsQrtPlanirovshik + " НЕ РАВНО!!!! " + totalPcsSpravochnikQrt +
                                      " (упаковки за 1-ый квартал)");
                }
                else
                {
                    Console.WriteLine(totalPcsQrtPlanirovshik + " = " + totalPcsSpravochnikQrt +
                                      " (упаковки за 1-ый квартал)");
                }


                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                Thread.Sleep(500);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.FindPreparationInputFieldXPath)));

            } //конец цикла FOR перебора всех препаратов

        } //конец метода

        public void CheckPreparationListForPM(int user)
        {
            _preparationNamePlanirovschik.Clear();

            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);

            for (int i = 1; i <= _numberTableRows; i++)
            {
                var name =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]"))
                        .Text.Trim()
                        .Replace("\u00A0", " ")
                        .ToLower();
                var rowData = new RowData()
                {
                    IdPrUniq = Convert.ToInt32(tableRows[i - 1].GetAttribute("data_id")), // add preparation id
                    Id_BU = Convert.ToInt32(tableRows[i - 1].GetAttribute("bu_id")), // add preparation bu_id
                    Name = Regex.Replace(name, @"\s+", " ") // add preparation name

                };
                _preparationNamePlanirovschik.Add(rowData);
            }

            var listPreparationIDSpravochnik = _preparationDataSpravochnik.GetIdListByUserWithoutAutoplan(user);
            var listPreparationIDPlanirovschik = _preparationNamePlanirovschik.GetIdList();

            var compareWebwithExcel = RowDataList.CompareNumbers(listPreparationIDPlanirovschik,
                listPreparationIDSpravochnik);
            var compareExcelWithWeb = RowDataList.CompareNumbers(listPreparationIDSpravochnik,
                listPreparationIDPlanirovschik);

            if (compareWebwithExcel.Count != 0)
            {
                Console.WriteLine("Препараты отсутствуют в справочнике");
                foreach (var d in compareWebwithExcel)
                {
                    Console.WriteLine(d);
                }
            }
            if (compareExcelWithWeb.Count != 0)
            {
                Console.WriteLine("Препараты отсутствуют в планировщике");
                foreach (var d in compareExcelWithWeb)
                {
                    Console.WriteLine(d);
                }
            }
            Console.WriteLine("User - " + user + ". Проверен.");
        }

        public void GetListPreparationFromExcelForUser(int user)
        {
            var listPreparationIDSpravochnik = _preparationDataSpravochnik.GetIdListByUserWithoutAutoplan(user);

            Console.WriteLine("User_" + user + "Препараты отсутствуют в планировщике:");
            foreach (var preparation in listPreparationIDSpravochnik)
            {
                Console.WriteLine(preparation);
            }
        }

        #endregion

        #region Compare Territorii

        public void StoreExcelDataTerritoriiPlanirovschik()
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt,
                @"D:\Sneghka\Instructions NEW\Planirovschik\Current_users_territory_1.xls", "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["FIO"].ToString(),
                    IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                    Position = row["Position"].ToString(),
                    Name1RegionName = row["DistrictName"].ToString()

                };
                _planirovschikTerritorii.Add(rowData);
            }
        }

        public void StoreExcelDataTerritoriiSpravochnik()
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt,
                @"D:\Sneghka\Instructions NEW\Planirovschik\Spravochnik_terr.xlsx", "Select * from [zone_of_resp$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["Full_name"].ToString(),
                    IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                    Position = row["Position"].ToString(),
                    Name1RegionName = row["DistrictName2"].ToString()

                };
                _spravochnikTerritorii.Add(rowData);
            }
        }

        public void CompareTerritoriiSpravochnikWithPlanirovschik()
        {
            _differenceSpravochikWithPlanirovschik = RowTerritoriiList.CompareTerritoriis(_spravochnikTerritorii,
                _planirovschikTerritorii);
            Console.WriteLine("Данные есть в Справочнике, но отсутствуют в планировщике");
            /* foreach (var x in _differenceSpravochikWithPlanirovschik)
                Console.WriteLine(x.Position + "/ "+ x.FIO + " /" + x.DistrictName);*/
        }

        public void CompareTerritoriiPlanirovschikWithSpravochnik()
        {
            _differencePlanirovschikWithSpravochik = RowTerritoriiList.CompareTerritoriis(_planirovschikTerritorii,
                _spravochnikTerritorii);
            var y = _differencePlanirovschikWithSpravochik.Count;
            Console.WriteLine("Данные есть в Планировщике, но отсутствуют в Справочнике");
            /* foreach (var z in _differencePlanirovschikWithSpravochik)
                 Console.WriteLine(z.Position + "/ " + z.FIO + " /" + z.DistrictName);*/

        }

        public void WriteDataToExcel()
        {
            Excel.Application myApp = new Excel.Application();
            myApp.Visible = true;


            Excel.Workbook wb = myApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            /* Excel.Workbook wb = myApp.Workbooks.Add(@"D:\Sneghka\Инструкции\7_week.xlsx");*/
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
            ws.Cells[1, 1] = "Есть в планировщике, но нет в справочнике";
            ws.Cells[2, 1] = "Position";
            ws.Cells[2, 2] = "Fio";
            ws.Cells[2, 3] = "DistrictName";

            for (int i = 3; i <= _differencePlanirovschikWithSpravochik.Count; i++)
            {

                ws.Cells[i, 1] = _differencePlanirovschikWithSpravochik[i - 3].Position;
                ws.Cells[i, 2] = _differencePlanirovschikWithSpravochik[i - 3].FIO;
                ws.Cells[i, 3] = _differencePlanirovschikWithSpravochik[i - 3].Name1RegionName;
            }

            ws.Cells[1, 5] = "Есть в справочнике, но нет в планировщике";
            ws.Cells[2, 5] = "Position";
            ws.Cells[2, 6] = "Fio";
            ws.Cells[2, 7] = "DistrictName";

            for (int i = 3; i <= _differenceSpravochikWithPlanirovschik.Count; i++)
            {

                ws.Cells[i, 5] = _differenceSpravochikWithPlanirovschik[i - 3].Position;
                ws.Cells[i, 6] = _differenceSpravochikWithPlanirovschik[i - 3].FIO;
                ws.Cells[i, 7] = _differenceSpravochikWithPlanirovschik[i - 3].Name1RegionName;
            }
            wb.SaveAs(@"D:Sneghka\CompareTerritorii.xlsx");
            wb.Close(Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            myApp.Quit();
        }

        #endregion

        #region Compare BuId

        public void StoreExcelDataBuIdPlanirovschik()
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt,
                @"D:\Sneghka\Instructions NEW\Planirovschik\Current_preparation_bu.xls", "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["FIO"].ToString(),
                    IdSotr = Convert.ToInt32(row["id"]),
                    BuId = Convert.ToInt32(row["roleName"])

                };
                planirovschikBuId.Add(rowData);
            }
            Console.WriteLine("Данные из Планировщика");
            foreach (var x in planirovschikBuId)
            {
                Console.WriteLine(x.FIO + " - " + x.BuId);
            }

        }

        public void StoreExcelDataBuIdSpravochnik()
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt,
                @"D:\Sneghka\Instructions NEW\Planirovschik\Spravochnik_bu.xlsx", "Select * from [zone_of_resp$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["Full_name"].ToString(),
                    IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                    BuId = Convert.ToInt32(row["BUID"])

                };
                spravochnikBuId.Add(rowData);
            }
            var y = spravochnikBuId.GetUniqueNoteBySeveralFields();
            Console.WriteLine("Данные из справочника");
            foreach (var x in y)
            {
                Console.WriteLine(x.IdSotr + " - " + x.FIO + " - " + x.BuId);
            }
        }

        public void CompareBuIdSpravochnikWithPlanirovschik()
        {
            Excel.Application myApp = new Excel.Application();
            myApp.Visible = true;


            Excel.Workbook wb = myApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            /* Excel.Workbook wb = myApp.Workbooks.Add(@"D:\Sneghka\Инструкции\7_week.xlsx");*/
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            var spravochnikUniq = spravochnikBuId.GetUniqueNoteBySeveralFields();
            var difference = RowTerritoriiList.CompareBuId(spravochnikUniq, planirovschikBuId);

            ws.Cells[1, 1] = "Есть в справочнике, но нет в планировщике";
            ws.Cells[2, 1] = "Id_Sotr";
            ws.Cells[2, 2] = "Fio";
            ws.Cells[2, 3] = "Bu_ID";

            for (int i = 3; i <= difference.Count; i++)
            {

                ws.Cells[i, 1] = difference[i - 3].IdSotr;
                ws.Cells[i, 2] = difference[i - 3].FIO;
                ws.Cells[i, 3] = difference[i - 3].BuId;
            }
            wb.SaveAs(@"D:Sneghka\CompareBuId_Sprav_Plan.xlsx");
            wb.Close(Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            myApp.Quit();

        }

        public void CompareBuIdPlanirovschikWithSpravochnik()
        {
            Excel.Application myApp = new Excel.Application();
            myApp.Visible = true;


            Excel.Workbook wb = myApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            /* Excel.Workbook wb = myApp.Workbooks.Add(@"D:\Sneghka\Инструкции\7_week.xlsx");*/
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            var spravochnikUniq = spravochnikBuId.GetUniqueNoteBySeveralFields();
            var difference = RowTerritoriiList.CompareBuId(planirovschikBuId, spravochnikUniq);
            var newDifference = new RowTerritoriiList();
            foreach (var z in difference)
            {
                if (!(z.BuId == 81 || z.BuId == 82))
                {
                    newDifference.Add(z);
                }
                else
                {
                    foreach (var c in difference)
                    {
                        if (c.IdSotr == z.IdSotr && c.BuId == 114) newDifference.Add(z);
                        else
                        {
                            continue;
                        }
                    }
                }
            }

            ws.Cells[1, 1] = "Есть в планировщике, но нет в справочнике";
            ws.Cells[2, 1] = "Id_Sotr";
            ws.Cells[2, 2] = "Fio";
            ws.Cells[2, 3] = "Bu_ID";

            for (int i = 3; i <= newDifference.Count; i++)
            {

                ws.Cells[i, 1] = newDifference[i - 3].IdSotr;
                ws.Cells[i, 2] = newDifference[i - 3].FIO;
                ws.Cells[i, 3] = newDifference[i - 3].BuId;
            }
            wb.SaveAs(@"D:Sneghka\CompareBuId_Plan_Sprav.xlsx");
            wb.Close(Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            myApp.Quit();

        }


        #endregion

        #region Distribution

        public void StoreExcelDataDistribution()
        {
            DataTable dt2015 = new DataTable();
            DataTable dt2016 = new DataTable();

            /*  WorkWithExcelFile.ExcelFileToDataTable(out dt2015, @"D:\Sneghka\Selenium\Projects\Planirovschik\Disrtibution_total.xlsx",
                  "Select * from [2015total$]");
              foreach (DataRow row in dt2015.Rows)
              {
                  if (row[0] == DBNull.Value) continue;

                  var rowData = new RowData
                  {
                      IdPrUniq = Convert.ToInt32(row["Id_2015"]),
                      Name = row["Препарат_2015"].ToString(),
                      Upakovki = Convert.ToInt32(row["pcs_2015"])
                  };
                  _distribution2015XlsList.Add(rowData);
              }
  */
            WorkWithExcelFile.ExcelFileToDataTable(out dt2016,
                @"D:\Sneghka\Selenium\Projects\Planirovschik\Disrtibution_total_2016_Moscow.xlsx",
                "Select * from [2016total$]");
            foreach (DataRow row in dt2016.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var rowData = new RowData
                {
                    IdPrUniq = Convert.ToInt32(row["Id_2016"]),
                    Name = row["Препарат_2016"].ToString(),
                    Upakovki = Convert.ToInt32(row["pcs_2016"])
                };
                _distribution2016XlsList.Add(rowData);
            }
            Console.WriteLine("Excel was stored");
        }

        public void StoreExcelDataDistributionFromSpravochnik(string path, string pageName)
        {

            DataTable dt2016 = new DataTable();

            WorkWithExcelFile.ExcelFileToDataTable(out dt2016, path, $"Select * from [{pageName}$]");
            DataColumnCollection columns = dt2016.Columns;

            foreach (DataRow row in dt2016.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var rowData = new DistributionSpravochnikRow
                {
                    PreparationId = Convert.ToInt32(row["Id_PrUniq"]),
                    PreparationName = columns.Contains("SKU") ? row["SKU"].ToString() : row["Препарат"].ToString(),
                    Year = Convert.ToInt32(row["Год"]),
                    Month = Convert.ToInt32(row["Месяц"]),
                    Upakovki =
                        columns.Contains("уп") ? Convert.ToInt32(row["уп"]) : Convert.ToInt32(row["Сумма, упаковки"]),
                    Segment = columns.Contains("Сегмент") ? Convert.ToInt32(row["Сегмент"]) : (int?)null,
                    RegionName1 = columns.Contains("Name_1") ? row["Name_1"].ToString() : row["Регион"].ToString(),
                    DistrictName2 = columns.Contains("Name_2") ? row["Name_2"].ToString() : row["Область"].ToString(),
                    SalesTypeId = columns.Contains("SalesTypeID") ? Convert.ToInt32(row["SalesTypeID"]) : (int?)null
                };
                _distributionSpravochnikRows.Add(rowData);
            }
            Console.WriteLine("Excel was stored - " + path);
        }

        public void CheckDistributionDataWithExcel()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(1000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Console.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {

                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButton =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";


                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }
                if (preparationId < 0) preparationId *= -1; //  change id from negetive value to positive value
                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);

                // Блок сбора данных за 2016 год

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2016Xpath)));
                Helper.TryToClickWithoutException(PageElements.SalesData2016Xpath, _firefox);
                Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2016Xpath, "class", "ui-tabs-selected",
                    _firefox);
                Thread.Sleep(200);
                var total2016 = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2016.Text.Replace(" ", ""));
                if (total2016 == _distribution2016XlsList.GetUpakovkiById(preparationId))
                {
                    Console.WriteLine(preparationName + "_2016 (web/xls): " + total2016 + " = " +
                                      _distribution2016XlsList.GetUpakovkiById(preparationId));
                }
                else
                {
                    Console.WriteLine(preparationName + "_2016 (web/xls): " + total2016 + " НЕ РАВНО!!!! " +
                                      _distribution2016XlsList.GetUpakovkiById(preparationId));
                }
                // Блок сбора данных за 2017 год (для проверки 12.01.17 этот блок не нужен - сверяем только 2016 год)

                /* Helper.TryToClickWithoutException(PageElements.SalesData2017Xpath, _firefox);
                 Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2017Xpath, "class", "ui-tabs-selected", _firefox);
                 Thread.Sleep(200);
                 var totalSum2017 = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2017.Text.Replace(" ", ""));
                 if (totalSum2017 == _distribution2016XlsList.GetUpakovkiById(preparationId))
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " = " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }
                 else
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " НЕ РАВНО!!!! " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }*/
                Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);


            } // end FOR loop
        }

        public void CheckDistributionDataWithExcelRussia()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(PageElements.PreparationTable)));
            Thread.Sleep(500);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Console.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(PageElements.PreparationTable)));
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]/td[3]")).Text;
                var raschetButton =
                    _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]/td[6]/input[1]"));
                var raschetButtonXPath = $".//*[@id='preparation_info']/tbody/tr[{i}]/td[6]/input[1]";

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("№" + i + " " + preparationName + " - кнопка Расчет неактивна");
                    continue;
                }

                if (preparationId < 0) preparationId *= -1; //  change id from negetive value to positive value
                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);

                // Блок сбора данных за 2016 год

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2016Xpath)));
                Helper.TryToClickWithoutException(PageElements.SalesData2016Xpath, _firefox);
                Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2016Xpath, "class", "ui-tabs-selected",
                    _firefox);
                Thread.Sleep(200);
                var total2016Web = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2016.Text.Replace(" ", ""));
                var total2016Xls = _distributionSpravochnikRows.GetUpakovkiByIdWithoutCrimea(preparationId);
                if (total2016Web == total2016Xls)
                {
                    Console.WriteLine("№" + i + " " + preparationName + "_2016 (web/xls): " + total2016Web + " = " +
                                      total2016Xls);
                }
                else
                {
                    Console.WriteLine("№" + i + " " + preparationName + "_2016 (web/xls): " + total2016Web +
                                      " НЕ РАВНО!!!! " + total2016Xls);
                }
                // Блок сбора данных за 2017 год (для проверки 12.01.17 этот блок не нужен - сверяем только 2016 год)

                /* Helper.TryToClickWithoutException(PageElements.SalesData2017Xpath, _firefox);
                 Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2017Xpath, "class", "ui-tabs-selected", _firefox);
                 Thread.Sleep(200);
                 var totalSum2017 = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2017.Text.Replace(" ", ""));
                 if (totalSum2017 == _distribution2016XlsList.GetUpakovkiById(preparationId))
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " = " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }
                 else
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " НЕ РАВНО!!!! " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }*/
                Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);


            } // end FOR loop
        }

        public void CheckDistributionDataWithExcelForNop()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(PageElements.PreparationTable)));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='plan_status_select']/label[1]/span")));
            // кнопка "Продуктов {0}" , число строк в таблице
            Thread.Sleep(500);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Console.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(PageElements.PreparationTable)));
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]"))
                            .GetAttribute("data_id"));
                var preparationBU =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]"))
                            .GetAttribute("bu_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]/td[3]")).Text;
                var raschetButton =
                    _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]/td[6]/input[1]"));
                var raschetButtonXPath = $".//*[@id='preparation_info']/tbody/tr[{i}]/td[6]/input[1]";


                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("№" + i + " " + preparationName + " - кнопка Расчет неактивна");
                    continue;
                }

                int total2016Xls;
                total2016Xls =
                    _distributionSpravochnikRows.GetUpakovkiByIdBySegmentBySalesTypeWithoutCrimea(preparationId);

                if (preparationBU == 67 || preparationBU == 115)
                {
                    total2016Xls = _distributionSpravochnikRows.GetUpakovkiByIdWithoutCrimea(preparationId);
                }
                if (preparationBU == 33)
                {
                    preparationId *= -1; //  change id from negetive value to positive value
                    total2016Xls =
                        _distributionSpravochnikRows.GetUpakovkiByIdBySegmentBySalesTypeWithoutCrimeaLgotaBU33(
                            preparationId);
                }

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);

                // Блок сбора данных за 2016 год

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2016Xpath)));
                Helper.TryToClickWithoutException(PageElements.SalesData2016Xpath, _firefox);
                Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2016Xpath, "class", "ui-tabs-selected",
                    _firefox);
                Thread.Sleep(200);
                var total2016Web = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2016.Text.Replace(" ", ""));

                if (total2016Web == total2016Xls)
                {
                    Console.WriteLine("№" + i + " " + preparationName + "_2016 (web/xls): " + total2016Web + " = " +
                                      total2016Xls);
                }
                else
                {
                    Console.WriteLine("№" + i + " " + preparationName + "_2016 (web/xls): " + total2016Web +
                                      " НЕ РАВНО!!!! " + total2016Xls);
                }
                // Блок сбора данных за 2017 год (для проверки 12.01.17 этот блок не нужен - сверяем только 2016 год)

                /* Helper.TryToClickWithoutException(PageElements.SalesData2017Xpath, _firefox);
                 Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2017Xpath, "class", "ui-tabs-selected", _firefox);
                 Thread.Sleep(200);
                 var totalSum2017 = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2017.Text.Replace(" ", ""));
                 if (totalSum2017 == _distribution2016XlsList.GetUpakovkiById(preparationId))
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " = " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }
                 else
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " НЕ РАВНО!!!! " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }*/
                Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);


            } // end FOR loop
        }

        public void CheckDistributionDataWithExcelForTm(string region)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(PageElements.PreparationTable)));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='plan_status_select']/label[1]/span")));
            // кнопка "Продуктов {0}" , число строк в таблице
            Thread.Sleep(500);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Console.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(PageElements.PreparationTable)));
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]/td[3]")).Text;
                var raschetButton =
                    _firefox.FindElement(By.XPath($".//*[@id='preparation_info']/tbody/tr[{i}]/td[6]/input[1]"));
                var raschetButtonXPath = $".//*[@id='preparation_info']/tbody/tr[{i}]/td[6]/input[1]";

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("№" + i + " " + preparationName + " - кнопка Расчет неактивна");
                    continue;
                }

                if (preparationId < 0) preparationId *= -1; //  change id from negetive value to positive value
                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);

                // Блок сбора данных за 2016 год

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2016Xpath)));
                Helper.TryToClickWithoutException(PageElements.SalesData2016Xpath, _firefox);
                Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2016Xpath, "class", "ui-tabs-selected",
                    _firefox);
                Thread.Sleep(200);
                var total2016Web = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2016.Text.Replace(" ", ""));

                var total2016Xls =
                    _distributionSpravochnikRows.GetUpakovkiByIdBySegmentBySalesTypeByRegion(preparationId, region);
                if (total2016Web == total2016Xls)
                {
                    Console.WriteLine("№" + i + " " + preparationName + "_2016 (web/xls): " + total2016Web + " = " +
                                      total2016Xls);
                }
                else
                {
                    Console.WriteLine("№" + i + " " + preparationName + "_2016 (web/xls): " + total2016Web +
                                      " НЕ РАВНО!!!! " + total2016Xls);
                }
                // Блок сбора данных за 2017 год (для проверки 12.01.17 этот блок не нужен - сверяем только 2016 год)

                /* Helper.TryToClickWithoutException(PageElements.SalesData2017Xpath, _firefox);
                 Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2017Xpath, "class", "ui-tabs-selected", _firefox);
                 Thread.Sleep(200);
                 var totalSum2017 = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2017.Text.Replace(" ", ""));
                 if (totalSum2017 == _distribution2016XlsList.GetUpakovkiById(preparationId))
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " = " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }
                 else
                 {
                     Console.WriteLine(preparationName + "_2017 (web/xls): " + totalSum2017 + " НЕ РАВНО!!!! " +
                                       _distribution2016XlsList.GetUpakovkiById(preparationId));
                 }*/
                Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);


            } // end FOR loop
        }

        #endregion

        #region Audit.

        public void LoginStadaAudit(string url, string login, string password)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            ((IJavaScriptExecutor)_firefox).ExecuteScript("window.open()");
            List<string> handles = _firefox.WindowHandles.ToList();
            _firefox.SwitchTo().Window(handles.Last());
            _planirovschikdWindow = _firefox.CurrentWindowHandle;
            _firefox.Navigate().GoToUrl(url);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
            pageElements.LoginField.SendKeys(login);
            pageElements.PasswordField.SendKeys(password);
            pageElements.SubmitButton.Click();
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='dialog_init']")));
            Thread.Sleep(2000);
        }


        public void LoginDashBoardAudit()
        {

            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            _firefox.Navigate()
                .GoToUrl(
                    "http://pharmxplorer.com.ua/QvAJAXZfc/opendoc.htm?document=TestDocs/stada/xls_data_test.qvw&host=QVS@qlikview&anonymous=true");
            Thread.Sleep(4000);

            var iframe = _firefox.FindElement(By.Id("popupFrame1"));
            _firefox.SwitchTo().Frame(iframe);
            _firefox.FindElement(By.XPath("html/body/div/table/tbody/tr[1]/td[2]/input")).SendKeys("TEST");
            Thread.Sleep(1000);
            _firefox.FindElement(By.XPath("html/body/div/table/tbody/tr[2]/td[2]/input")).SendKeys("WEQCCR@#$FE");
            _firefox.FindElement(By.XPath(".//*[@id='PageFooter']/td/button[1]")).Click();
            Thread.Sleep(3000);
            _firefox.SwitchTo().DefaultContent();
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='MainContainer']")));
            _dashBoardWindow = _firefox.CurrentWindowHandle;
            Thread.Sleep(2000);
        }

        public void SetUpFilterDashBoardAudit()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.LockAuditWebXPath)));
            Helper.TryToClickWithoutException(PageElements.LockAuditWebXPath, _firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchPeriodAuditWebXPath)));
            Helper.TryToClickWithoutException(PageElements.SearchPeriodAuditWebXPath, _firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.InputFieldAuditXPath)));
            pageElements.InputFieldAuditWeb.SendKeys("2015" + Keys.Enter);

            Helper.TryToClickWithoutException(PageElements.AreaLevel_2AuditWebXPath, _firefox);

            Waiting.WaitForAjax(_firefox);
            Thread.Sleep(1000);
            Helper.TryToClickWithoutException(PageElements.SearchAreaNameAuditWebXPath, _firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.InputFieldAuditXPath)));
            pageElements.InputFieldAuditWeb.SendKeys("Россия" + Keys.Enter);

            Waiting.WaitForAjax(_firefox);
            Thread.Sleep(1000);
        }

        public void StoreExcelDataAudit()
        {
            DataTable dtAudit2015 = new DataTable();


            WorkWithExcelFile.ExcelFileToDataTable(out dtAudit2015,
                @"D:\Sneghka\Selenium\Projects\Planirovschik\DataForCheck_20161220.xlsx",
                "Select * from [Audit_2016$]");
            foreach (DataRow row in dtAudit2015.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var rowData = new RowData
                {
                    IdPrUniq = Convert.ToInt32(row["preparationId"]),
                    Name = row["name"].ToString(),
                    Upakovki = Convert.ToInt32(row["Свои упаковки"]),
                    UpakovkiConcurent = Convert.ToInt32(row["Конкурентные упаковки"]),
                    AreaName = row["area_name"].ToString(),
                    SalesType = Convert.ToInt32(row["sales_type"])
                };
                _audit2015XlsList.Add(rowData);
            }
            Console.WriteLine("Page Audit2015 was stored");

        }

        public void CheckAuditDataWithExcel()
        {
            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXpath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";

                if (preparationId < 0)
                {
                    preparationId *= -1;
                }
                if (
                    !_firefox.FindElement(By.XPath(raschetButtonXpath))
                        .GetAttribute("class")
                        .Contains("ui-button-disabled")) // начало проверки на активность кнопки РАСЧЁТ  
                {
                    ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);",
                        _firefox.FindElement(By.XPath(raschetButtonXpath)));
                    Thread.Sleep(500);
                    Helper.TryToClickWithoutException(raschetButtonXpath, _firefox);
                    // click "Расчёт" для выбранного элемента
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AuditDataOwn2016XPath)));
                    Helper.TryToClickWithoutException(PageElements.AuditDataOwn2016XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataOwn2016XPath, "class",
                        "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);
                    var totalOwn2015 = Convert.ToInt32(pageElements.TotalSumOwnSalesData2015.Text.Replace(" ", ""));

                    if (totalOwn2015 == _audit2015XlsList.GetUpakovkiById(preparationId))
                    {
                        Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " + totalOwn2015 +
                                          " = " + _audit2015XlsList.GetUpakovkiById(preparationId));
                    }
                    else
                    {
                        Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " + totalOwn2015 +
                                          " НЕ РАВНО!!! " + _audit2015XlsList.GetUpakovkiById(preparationId));
                    }

                    Helper.TryToClickWithoutException(PageElements.AuditDataCompetitor2016XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataCompetitor2016XPath, "class",
                        "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);
                    var totalCompetitior2015 =
                        Convert.ToInt32(pageElements.TotalSumCompetitorSalesData2015.Text.Replace(" ", ""));

                    if (totalCompetitior2015 == _audit2015XlsList.GetUpakovkiConcurentById(preparationId))
                    {
                        Console.WriteLine(preparationId + " " + preparationName + "_2016Competitor (web/xls): " +
                                          totalCompetitior2015 + " = " +
                                          _audit2015XlsList.GetUpakovkiConcurentById(preparationId));
                    }
                    else
                    {
                        Console.WriteLine(preparationId + " " + preparationName + "_2016Competitor (web/xls): " +
                                          totalCompetitior2015 + " НЕ РАВНО!!!! " +
                                          _audit2015XlsList.GetUpakovkiConcurentById(preparationId));
                    }
                    Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);

                } // end IF check button
                else
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }

            }
        }

        public void CheckAuditDataWithExcelForPm()
        {
            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXpath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";

                if (preparationId < 0)
                {
                    preparationId *= -1;
                }
                if (
                    !_firefox.FindElement(By.XPath(raschetButtonXpath))
                        .GetAttribute("class")
                        .Contains("ui-button-disabled")) // начало проверки на активность кнопки РАСЧЁТ  
                {
                    ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);",
                        _firefox.FindElement(By.XPath(raschetButtonXpath)));
                    Thread.Sleep(500);
                    Helper.TryToClickWithoutException(raschetButtonXpath, _firefox);
                    // click "Расчёт" для выбранного элемента
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AuditDataOwn2016XPath)));
                    Helper.TryToClickWithoutException(PageElements.AuditDataOwn2016XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataOwn2016XPath, "class",
                        "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);
                    var totalOwn2015 = Convert.ToInt32(pageElements.TotalSumOwnSalesData2015.Text.Replace(" ", ""));
                    var totalOwnSpravochnikUpakovki = _audit2015XlsList.GetUpakovkiByIdAndSalesType(preparationId, 1);

                    if (totalOwn2015 == totalOwnSpravochnikUpakovki)
                    {
                        Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " + totalOwn2015 +
                                          " = " + totalOwnSpravochnikUpakovki);
                    }
                    else
                    {

                        if (preparationName.Contains("ЛЬГОТА"))
                        {
                            if (totalOwn2015 == _audit2015XlsList.GetUpakovkiByIdAndTwoSalesType(preparationId, 2, 3))
                            {
                                Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                  totalOwn2015 + " = " +
                                                  _audit2015XlsList.GetUpakovkiByIdAndTwoSalesType(preparationId, 2, 3) +
                                                  " (ЛЬГОТА + ТЕНДЕР)");
                            }
                            else
                            {
                                Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                  totalOwn2015 + " НЕ РАВНО!!! " +
                                                  _audit2015XlsList.GetUpakovkiByIdAndTwoSalesType(preparationId, 2, 3) +
                                                  " (ЛЬГОТА + ТЕНДЕР)");
                            }
                        }
                        else
                        {
                            if (totalOwnSpravochnikUpakovki / totalOwn2015 == 2)
                            {
                                Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                  totalOwn2015 + " = " + totalOwnSpravochnikUpakovki / 2);
                            }
                            else
                            {


                                if (totalOwn2015 ==
                                    _audit2015XlsList.GetUpakovkiByIdAndTwoSalesType(preparationId, 1, 3))
                                {
                                    Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                      totalOwn2015 + " = " +
                                                      _audit2015XlsList.GetUpakovkiByIdAndTwoSalesType(preparationId, 1,
                                                          3) + " (КОММЕРЦИЯ + ТЕНДЕР)");
                                }
                                else
                                {
                                    Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                      totalOwn2015 + " НЕ РАВНО!!! " +
                                                      _audit2015XlsList.GetUpakovkiByIdAndTwoSalesType(preparationId, 1,
                                                          3) + " (КОММЕРЦИЯ + ТЕНДЕР)");
                                }
                            }

                        }
                    }

                    Helper.TryToClickWithoutException(PageElements.AuditDataCompetitor2016XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataCompetitor2016XPath, "class",
                        "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);
                    var totalCompetitior2015 =
                        Convert.ToInt32(pageElements.TotalSumCompetitorSalesData2015.Text.Replace(" ", ""));
                    var totalCompetitorSpravochikpakovki =
                        _audit2015XlsList.GetUpakovkiConcurentByIdAndSalesType(preparationId, 1);

                    if (totalCompetitior2015 == totalCompetitorSpravochikpakovki)
                    {
                        Console.WriteLine(preparationId + " " + preparationName + "_2016Competitor (web/xls): " +
                                          totalCompetitior2015 + " = " + totalCompetitorSpravochikpakovki);
                    }
                    else
                    {

                        if (preparationName.Contains("ЛЬГОТА"))
                        {
                            if (totalCompetitior2015 ==
                                _audit2015XlsList.GetUpakovkiConcurentByIdAndTwoSalesType(preparationId, 2, 3))
                            {
                                Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                  totalCompetitior2015 + " = " +
                                                  _audit2015XlsList.GetUpakovkiConcurentByIdAndTwoSalesType(
                                                      preparationId, 2, 3));
                            }
                            else
                            {
                                Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                  totalCompetitior2015 + " НЕ РАВНО!!! " +
                                                  _audit2015XlsList.GetUpakovkiConcurentByIdAndTwoSalesType(
                                                      preparationId, 2, 3));
                            }

                        }
                        else
                        {
                            if (totalCompetitorSpravochikpakovki / totalCompetitior2015 == 2)
                            {
                                Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                  totalCompetitior2015 + " = " + totalCompetitorSpravochikpakovki / 2);

                            }
                            else
                            {
                                if (totalCompetitior2015 ==
                                    _audit2015XlsList.GetUpakovkiConcurentByIdAndTwoSalesType(preparationId, 1, 3))
                                {
                                    Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                      totalCompetitior2015 + " = " +
                                                      _audit2015XlsList.GetUpakovkiConcurentByIdAndTwoSalesType(
                                                          preparationId, 1, 3));
                                }
                                else
                                {
                                    Console.WriteLine(preparationId + " " + preparationName + "_2016Own (web/xls): " +
                                                      totalCompetitior2015 + " НЕ РАВНО!!! " +
                                                      _audit2015XlsList.GetUpakovkiConcurentByIdAndTwoSalesType(
                                                          preparationId, 1, 3));
                                }
                            }

                        }
                    }
                    Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);

                } // end IF check button
                else
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }

            }
        }

        public void CheckAuditDataWithDashBoard()
        {
            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);

            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[2]")).Text;

                if (
                    !_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]"))
                        .GetAttribute("class")
                        .Contains("ui-button-disabled")) // начало проверки на активность кнопки РАСЧЁТ  
                {

                    ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);",
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]")));
                    Thread.Sleep(500);
                    Helper.TryToClickWithoutException(
                        ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]", _firefox);
                    // click "Расчёт" для выбранного элемента
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2016Xpath)));
                    Helper.TryToClickWithoutException(PageElements.AuditDataOwn2016XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataOwn2016XPath, "class",
                        "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);

                    var totalOwn2015Plan = Convert.ToInt32(pageElements.TotalSumOwnSalesData2015.Text.Replace(" ", ""));

                    Helper.TryToClickWithoutException(PageElements.AuditDataCompetitor2016XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataCompetitor2016XPath, "class",
                        "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);

                    var totalCompetitior2015Plan =
                        Convert.ToInt32(pageElements.TotalSumCompetitorSalesData2015.Text.Replace(" ", ""));

                    Thread.Sleep(2000);

                    _firefox.SwitchTo().Window(_dashBoardWindow);

                    wait.Until(
                        ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchPreperationIdAuditWebXPath)));

                    Helper.TryToClickWithoutException(PageElements.SearchPreperationIdAuditWebXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.InputFieldAuditXPath)));
                    pageElements.InputFieldAuditWeb.SendKeys("(" + preparationId + ")" + Keys.Enter);
                    Waiting.WaitPatternPresentInAttribute(".//*[@id='57']/div[3]/div/div[1]/div[1]", "class",
                        "QvSelected", _firefox);
                    Thread.Sleep(2000);
                    var totalOwn2015Dash = Convert.ToInt32(pageElements.TotalOwnPcsAuditWeb.GetAttribute("title"));
                    var totalCompetitor2015Dash =
                        Convert.ToInt32(pageElements.TotalCompetitorPcsAuditWeb.GetAttribute("title"));

                    if (totalOwn2015Plan == totalOwn2015Dash)
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " + totalOwn2015Plan +
                                          " = " + totalOwn2015Dash);
                    }
                    else
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " + totalOwn2015Plan +
                                          " НЕ РАВНО!!! " + totalOwn2015Dash);
                    }


                    if (totalCompetitior2015Plan == totalCompetitor2015Dash)
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " +
                                          totalCompetitior2015Plan + " = " + totalCompetitor2015Dash);
                    }
                    else
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " +
                                          totalCompetitior2015Plan + " НЕ РАВНО!!!! " + totalCompetitor2015Dash);
                    }
                    _firefox.SwitchTo().Window(_planirovschikdWindow);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.RaschetPlanaButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                } // end IF check Расчет button
                else
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }
            }

        }


        #endregion

        #region Chains

        public void LogoutStada(string url)
        {
            _firefox.Navigate().GoToUrl(url);
        }

        public void StoreLoginPasswordFromExcel()
        {
            DataTable dtLoginPassword = new DataTable();

            WorkWithExcelFile.ExcelFileToDataTable(out dtLoginPassword,
                @"D:\Sneghka\Selenium\Projects\Planirovschik\Check_Login_Pass.xlsx",
                "Select * from [Пользователи$]");
            foreach (DataRow row in dtLoginPassword.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var loginPassword = new LoginPassword()
                {
                    Login = row["login"].ToString(),
                    Password = row["parole"].ToString()
                };
                _loginPasswordList.Add(loginPassword);
            }
            Console.WriteLine("Login/Password was stored");
        }

        public void CheckLoginPasswordMethod1(string url)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            _firefox.Navigate().GoToUrl(url);
            int i = 1;
            for (int a = 189; a < _loginPasswordList.Count; a++)
            {

                pageElements.LoginField.Clear();
                pageElements.PasswordField.Clear();
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
                pageElements.LoginField.SendKeys(_loginPasswordList[a].Login);
                pageElements.PasswordField.SendKeys(_loginPasswordList[a].Password);
                pageElements.SubmitButton.Click();

                Thread.Sleep(4000);
                if (_firefox.FindElement(By.XPath(".//*[@id='dialog-confirm']")).GetAttribute("style") ==
                    "display: none;")
                {
                    if (Helper.IsElementPresent(By.XPath("html/body/div[4]/div[3]/div/button[1]"), _firefox))
                    {
                        _firefox.FindElement(By.XPath("html/body/div[4]/div[3]/div/button[1]")).Click();
                        Thread.Sleep(100);
                        _firefox.FindElement(By.XPath(".//*[@id='logoutButton']/a")).Click();
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("html/body/div[4]/div[3]/div/button[1]")));
                        Thread.Sleep(100);
                        _firefox.FindElement(By.XPath("html/body/div[4]/div[3]/div/button[1]")).Click();
                        Thread.Sleep(100);
                    }
                    else
                    {
                        _firefox.FindElement(By.XPath(".//*[@id='logoutButton']/a")).Click();
                        Thread.Sleep(100);
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("html/body/div[2]/div[3]/div/button[1]")));
                        Thread.Sleep(100);
                        _firefox.FindElement(By.XPath("html/body/div[2]/div[3]/div/button[1]")).Click();
                        Thread.Sleep(100);

                    }

                    Console.WriteLine("№" + i + "  Ok: " + _loginPasswordList[a].Login + " / " +
                                      _loginPasswordList[a].Password);
                    i++;
                }

                else
                {
                    Console.WriteLine("№" + i + "  Incorrect login or password: " + _loginPasswordList[a].Login + " / " +
                                      _loginPasswordList[a].Password);
                    _firefox.Navigate().GoToUrl(url);
                    Thread.Sleep(200);
                    i++;
                    continue;
                }


            }
        }

        public void CheckLoginPasswordMethod2(string url, string logout)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            _firefox.Navigate().GoToUrl(url);
            int i = 1;
            for (int a = 0; a < _loginPasswordList.Count; a++)
            {
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
                pageElements.LoginField.SendKeys(_loginPasswordList[a].Login);
                pageElements.PasswordField.SendKeys(_loginPasswordList[a].Password);
                pageElements.SubmitButton.Click();
                //Waiting.WaitForAjax()();
                Waiting.WaitForAjax(_firefox);
                if (_firefox.FindElement(By.XPath(".//*[@id='dialog-confirm']")).GetAttribute("style") ==
                    "display: none;")
                {
                    Console.WriteLine("№" + i + "  Ok: " + _loginPasswordList[a].Login + " / " +
                                      _loginPasswordList[a].Password);
                    i++;
                }
                else
                {
                    Console.WriteLine("№" + i + "  Incorrect login or password: " + _loginPasswordList[a].Login + " / " +
                                      _loginPasswordList[a].Password);
                    i++;
                    //continue;
                }
                _firefox.Navigate().GoToUrl(logout);

            }
        }

        public void ChainsAccept()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(3000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[2]")).Text;
                var preparationBu =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("bu_id"));
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton = _firefox.FindElement(By.XPath(raschetButtonXPath));
                var raschetButtonPlanStatus = raschetButton.GetAttribute("plan_status");

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }
                if (raschetButtonPlanStatus != "")
                {
                    Console.WriteLine(preparationName + " - статус - " + raschetButtonPlanStatus);
                    continue;
                }

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);",
                    _firefox.FindElement(By.XPath(raschetButtonXPath)));
                Thread.Sleep(1500);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(raschetButtonXPath)));
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SavePlanButtonXPath)));
                if (_firefox.FindElement(By.XPath(".//*[@id='save_plan_customer']")).GetAttribute("aria-disabled") ==
                    "true")
                {
                    Console.WriteLine("№" + i + " " + preparationName +
                                      " - препарат НЕ утверждён. Кнопка 'Сохранить план' неактивна");
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='dialog_init']")));
                    continue;
                }

                Helper.TryToClickWithoutException(PageElements.SavePlanButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AcceptButtonXpath)));
                Helper.TryToClickWithoutException(PageElements.AcceptButtonXpath, _firefox);
                Thread.Sleep(200);

                _firefox.FindElement(By.XPath("/html/body/div[@class='ui-pnotify ']/div/div[4]/center/input")).Click();
                // click "Перейти к утверждению"
                Waiting.WaitForAjax(_firefox);

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AcceptPlanButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.AcceptPlanButtonXPath, _firefox);
                Waiting.WaitForAjax(_firefox);

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ConfirmPlanButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ConfirmPlanButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='decline_plan']")));

                Console.WriteLine("№" + i + " " + preparationName + " - препарат утверждён");

                Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='dialog_init']")));
            }
        }

        public void ChainsApprove()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            Thread.Sleep(2000);

            _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[1]/td[6]/input[2]")).Click();
            // First element's PlanButton 
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='plan_list']")));
            var tableRowsPlansApprove = _firefox.FindElements(By.XPath(".//*[@id='preparation_info_short']/tbody/tr"));
            // get list of preparation
            var numberTableRowsPlansApprove = tableRowsPlansApprove.Count;
            Thread.Sleep(4000);
            Debug.WriteLine(numberTableRowsPlansApprove + " кол-во строк в таблице");

            for (int i = 1; i <= numberTableRowsPlansApprove; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info_short']")));
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info_short']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationNameXPath = ".//*[@id='preparation_info_short']/tbody/tr[" + i + "]/td[2]";
                var preparationName = _firefox.FindElement(By.XPath(preparationNameXPath)).Text;
                var preparationStatusXPath = ".//*[@id='preparation_info_short']/tbody/tr[" + i + "]/td[3]";
                var preparationStatus = _firefox.FindElement(By.XPath(preparationStatusXPath));

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);",
                    _firefox.FindElement(By.XPath(preparationNameXPath)));
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(preparationNameXPath, _firefox); // click выбранный элемент

                if (preparationStatus.Text == "одобрен")
                {
                    Console.WriteLine("№" + i + " " + preparationId + " " + preparationName + " - препарат уже одобрен");
                    continue;
                }

                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(PageElements.ApprovePlanButtonXPath)));
                Helper.TryToClickWithoutException(PageElements.ApprovePlanButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("html/body/div[5]/div[3]/div/button[1]")));
                Helper.TryToClickWithoutException("html/body/div[5]/div[3]/div/button[1]", _firefox);
                Waiting.WaitForAjax(_firefox);
                Helper.TryToClickWithoutException(PageElements.RefreshButtonXPath, _firefox);
                Waiting.WaitPatternPresentInText(preparationStatusXPath, "одобрен", _firefox);
                Thread.Sleep(200);
                Console.WriteLine("№" + i + " " + preparationId + " " + preparationName + " - препарат одобрен");
            }
        }


        #endregion

        #region GR

        public void StoreGr()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);


            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            //Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= _numberTableRows; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButton =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }
                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                // click "Расчёт" для выбранного элемента
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Thread.Sleep(500);
                var valueGr =
                    _firefox.ExecuteScript("return document.getElementById('COMP_GR').previousSibling.innerHTML;");
                // instead innerHTML can use innerContent
                _grListValue.Add("№" + i + " " + preparationId + " " + preparationName + " -- " + valueGr);
                Console.WriteLine("№" + i + " " + preparationId + " " + preparationName + " -- " + valueGr);
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            }
            File.WriteAllLines(@"D:\Sneghka\Selenium\Created documents\GR_Value.doc", _grListValue);
        }

        public void PrintGR()
        {
            File.WriteAllLines(@"D:\Sneghka\GR_Value.doc", _grListValue);
        }

        public void IsGrUnchangeable()
        {

            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= _numberTableRows; i++)
            {

                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                if (!_preparationDataSpravochnik.GetIdList().Contains(preparationId)) continue;

                var preparationBuId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("bu_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));


                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("Кнопка расчёт неактивна - " + preparationName);
                    continue;
                }
                Thread.Sleep(1000);
                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TotalPcsXPath)));
                var sliderGrClass = pageElements.GrSlider.GetAttribute("class");

                if (!sliderGrClass.Contains("ui-slider-disabled"))
                {
                    Console.WriteLine("№" + i + " " + preparationId + " " + " " + preparationName + ": Слайдер АКТИВЕН");
                }
                else
                {
                    Console.WriteLine("№" + i + " " + preparationId + " " + " " + preparationName +
                                      ": Слайдер НЕАКТИВЕН");
                }

                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            }

        }

        public void IsGrChangeable()
        {

            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            _numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(_numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= _numberTableRows; i++)
            {

                var preparationId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("data_id"));
                if (!_preparationDataSpravochnik.GetIdList().Contains(preparationId)) continue;

                var preparationBuId =
                    Convert.ToInt32(
                        _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]"))
                            .GetAttribute("bu_id"));
                var preparationName =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton =
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));


                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine("Кнопка расчёт неактивна - " + preparationName);
                    continue;
                }
                Thread.Sleep(1000);
                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TotalPcsXPath)));
                var sliderGrClass = pageElements.GrSlider.GetAttribute("class");

                if (!sliderGrClass.Contains("ui-slider-disabled"))
                {
                    Console.WriteLine("№" + i + " " + preparationId + " " + " " + preparationName + ": Слайдер АКТИВЕН");
                }
                else
                {
                    Console.WriteLine("№" + i + " " + preparationId + " " + " " + preparationName +
                                      ": Слайдер НЕАКТИВЕН");
                }

                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            }

        }

        public string GetPasswordByUser(string user)
        {
            var password = (from r in _loginPasswordList
                            where r.Login == "user_" + user
                            select r.Password).ToList();
            return password[0].ToString();
        }

        #endregion


        #region CheckStadaPlans

        public void StoreExcelDataTerritoriiSpravochnik(string sheetName)
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt,
                @"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник3_14.12.16___custom_.xls", "Select * from [" + sheetName + "$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var buID = row["BUID"].ToString();

                if (!buID.Contains(','))
                {
                    var rowData = new RowTerritorii()
                    {
                        FIO = row["Full_name"].ToString(),
                        IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                        Position = row["Position"].ToString(),
                        BuId = Convert.ToInt32(buID),
                        Name1RegionName = row["District#Name3"].ToString(),
                        Name1RegionId = row["District#id3"].ToString(),
                        Name2OblastName = row["District#Name2"].ToString(),
                        Name2OblastId = row["District#id2"].ToString(),
                        Name3RaionName = row["District#Name1"].ToString(),
                        Name3RaionId = row["District#id1"].ToString()
                    };
                    _spravochnikTerritorii.Add(rowData);
                }

                if (buID.Contains(','))
                {
                    var buArray = buID.Split(',');
                    for (int i = 0; i < buArray.Length; i++)
                    {
                        var rowData = new RowTerritorii()
                        {
                            FIO = row["Full_name"].ToString(),
                            IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                            Position = row["Position"].ToString(),
                            BuId = Convert.ToInt32(buArray[i]),
                            Name1RegionName = row["District.Name3"].ToString(),
                            Name1RegionId = row["District.id3"].ToString(),
                            //Name1RegionId = Convert.ToInt32(row["id региона"]),
                            Name2OblastName = row["District.Name2"].ToString(),
                            Name2OblastId = row["District.id2"].ToString(),
                            Name3RaionName = row["District.Name1"].ToString(),
                            Name3RaionId = row["District.id1"].ToString()
                        };
                        _spravochnikTerritorii.Add(rowData);
                    }
                }
            }
        }

        public void StoreExcelDataEmailSpravochik(string sheetName)
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt,
                @"D:\Sneghka\Selenium\Projects\Planirovschik\1_для_модуля__Справочник3_14.12.16___custom_.xls", "Select * from [" + sheetName + "$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new User
                {
                    UserName = row["Full_name"].ToString(),
                    UserId = Convert.ToInt32(row["Id_Sotr"]),
                    Email = row["Mail"].ToString(),

                };
                _usersListForEmailSpravochnik.Add(rowData);
            }
        }

        public void GoToOdobreniePlanovTab()
        {
            var wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ClosePreparationListButtonXpath)));
            Helper.TryToClickWithoutException(PageElements.ClosePreparationListButtonXpath, _firefox);
            Waiting.WaitForAjax(_firefox);
            Helper.TryToClickWithoutException(PageElements.TopMenuOdobreniePlanovButtonXpath, _firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TableOdobrenieXpath)));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='dep_info']/tbody/tr[1]/td")));
        }

        public void CalculateAllPlans(string url) // метод на заглушках - ПРОВЕРЯТЬ!!!!
        {
            var wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var rowList = _firefox.FindElements(By.XPath(PageElements.TableOdobrenieRowsXpath));

            for (int i = 0; i < rowList.Count; i++)
            {
                var status = _firefox.FindElement(By.XPath($".//*[@id='dep_info']/tbody/tr[3]/td[{i}]"));
                var buNumder = _firefox.FindElement(By.XPath($".//*[@id='dep_info']/tbody/tr[{i}]/td[1]"));
                var approveTableButton = $".//*[@id='dep_info']/tbody/tr[{i}]/td[4]/input";

                try
                {
                    if (status.Text != "Готов для одобрения")
                    {
                        Console.WriteLine("БЮ НЕ ОДОБРЕН - " + buNumder + " статус - " + status.Text);
                        continue;
                    }
                    if (rowList[i].Text == "Одобрение") continue;
                    if (rowList[i].Text == "Рассылка") return;
                    Helper.TryToClickWithoutException(approveTableButton, _firefox);
                    wait.Until(
                        ExpectedConditions.ElementToBeClickable(By.XPath(PageElements.GlobalApprovePlanButton1340Xpath)));
                    Helper.TryToClickWithoutException(PageElements.GlobalApprovePlanButton1340Xpath, _firefox);
                    var waitPlanCalculating = new WebDriverWait(_firefox, TimeSpan.FromMinutes(30));
                    waitPlanCalculating.Until(
                        ExpectedConditions.InvisibilityOfElementLocated(
                            By.XPath(PageElements.GlobalApprovePlanButton1340Xpath)));
                    Helper.TryToClickWithoutException("html/body/div[7]/div[11]/div/button[1]", _firefox);
                    //Close Button
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='dep_info']/tbody/tr[1]/td")));
                    // ЗАголовок ОДОБРЕНИЕ 
                    Console.WriteLine("Одобрено: БЮ - " + buNumder);
                }
                catch (Exception e)
                {
                    Console.WriteLine("БЮ НЕ ОДОБРЕН (exception)  - " + buNumder);
                    Console.WriteLine("Exception - " + e.Message);
                    Console.WriteLine("Exception inner - " + e.InnerException);
                    _firefox.Navigate().GoToUrl(url);
                    GoToOdobreniePlanovTab();
                }
            }
        }

        public List<PlanTableRow> ReadPlanTable1340(int userId, int month) // получить коллекцию tr
        {
            var plan = new List<PlanTableRow>();
            var wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath($".//*[@id='plan_{userId}']/tbody")));
            Thread.Sleep(3000);

            var tableStructure = _firefox.FindElement(By.XPath($".//*[@id='plan_{userId}']/tbody")).GetAttribute("innerHTML");
            var userName = _firefox.FindElement(By.XPath(".//*[@id='msg-dialog-body']/p[2]")).Text;

            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(tableStructure);
            var tbRows = htmlDoc.DocumentNode.Descendants("tr");

            var preparationName1 = string.Empty;
            int i = 1;
            foreach (var tr in tbRows)
            {
                if (tr.InnerText == "&nbsp;" || tr.InnerText == " " || tr.InnerText == "")
                {
                    i++;
                    continue;
                }
                if (tr.InnerText.Contains("Total")) break;

                var childList = tr.ChildNodes;

                if (childList[0].Name == "th")
                {
                    preparationName1 = childList[0].InnerText;
                    i++;
                    continue;
                }

                for (int k = 1; k <= month; k++)
                {
                    var planTableRow = new PlanTableRow()
                    {
                        UserName = userName,
                        UserId = userId,
                        PreparationName = preparationName1,
                        TerritoriaName = childList[0].InnerText,
                        Month = k,
                        Pcs = Convert.ToInt32(Regex.Replace(childList[k].InnerText, @"\s+", ""))
                    };
                    plan.Add(planTableRow);
                    i++;
                }
            }
            return plan;
        }

        public List<PlanTableRow> ReadPlanTableOrdinaryUser(string url, int idUser, int month, FirefoxDriver firefox2, string logout)
        {
            var planForOrdinaryUser = new List<PlanTableRow>();
            //var firefox2 = new FirefoxDriver(SetUpFirefoxProfile());

            WebDriverWait wait2 = new WebDriverWait(firefox2, TimeSpan.FromSeconds(120));
            var action = new Actions(firefox2);
            var pageElements = new PageElements(firefox2);
            firefox2.Navigate().GoToUrl(url);
            wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
            pageElements.LoginField.SendKeys("user_" + idUser);
            pageElements.PasswordField.SendKeys("1");
            pageElements.SubmitButton.Click();
            Thread.Sleep(4000);

            if (!Helper.IsElementPresent(By.XPath("html/body/div[4]/div[3]/div/button[1]"), firefox2))
            {
                firefox2.Navigate().GoToUrl(logout);
                return planForOrdinaryUser;
            }

            wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath("html/body/div[4]/div[3]/div/button[1]"))); // CLOSE BUTTON
            var buNumber =firefox2.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[1]")).GetAttribute("bu_id");
            Helper.TryToClickWithoutException("html/body/div[4]/div[3]/div/button[1]", firefox2);
            wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElementsAdditional.TopMenuPlanyPoTerritoriamButton)));
            Helper.TryToClickWithoutException(PageElementsAdditional.TopMenuPlanyPoTerritoriamButton, firefox2);
            wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='regions_info_short']/tbody"))); // ТАБЛИЦА ТЕРРИТОРИЙ
            Thread.Sleep(2000);
            var terrList = firefox2.FindElements(By.XPath(".//*[@id='regions_info_short']/tbody/tr"));
            
            for (int j = 0; j < terrList.Count; j++)
            {
                var regionName = terrList[j].Text.Split(' ');

                if (regionsString.Contains(regionName[1])) continue;

                Helper.TryToClickWithoutException($".//*[@id='regions_info_short']/tbody/tr[{j + 1}]", firefox2);
                wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='plan']/tbody/tr[1]/th[2]")));
                Thread.Sleep(1000);
                var terrName = firefox2.FindElement(By.XPath($".//*[@id='regions_info_short']/tbody/tr[{j + 1}]/td[2]")).Text;
                // Заголовок "ПРЕПАРАТ"

                #region МЕТОД СЧИТЫВАНИЯ В ЕКСЕЛЬ
                /* Helper.TryToClickWithoutException(".//*[@id='export_plan_to_xls']", firefox2);
                 wait2.Until(ExpectedConditions.ElementIsVisible(By.XPath("html/body/div[5]/div[3]/div/button[1]")));
                 // Кнопка "ПРИНЯТЬ"
                 Helper.TryToClickWithoutException("html/body/div[5]/div[3]/div/button[1]", firefox2);
                 //Экспортируем файл в ексель
                 Thread.Sleep(5000);

                 var directory = new DirectoryInfo(@"D:\DownloadTest");
                 var myFile = (from f in directory.GetFiles()
                               orderby f.LastWriteTime descending
                               select f).First();

                 Console.WriteLine(myFile.Name);

                 DataTable dt = new DataTable();
                 WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\DownloadTest\" + myFile,
                     "SELECT * from [Worksheet$B2:F]");

                 foreach (DataRow row in dt.Rows)
                 {
                     if (row[0] == DBNull.Value) continue;
                     if (row["Препарат"].ToString() == "Итого") continue;

                     var rowData = new PlanTableRow()
                     {
                         UserId = idUser,
                         PreparationName = row["Препарат"].ToString(),
                         JanPsc = Convert.ToInt32(row["Январь"]),
                         FebPsc = Convert.ToInt32(row["Февраль"]),
                         MarPsc = Convert.ToInt32(row["Март"])
                     };
                     planForOrdinaryUser.Add(rowData);
                     //Console.WriteLine(rowData.UserId + " " + "Jan - " + rowData.JanPsc + "/ Feb -  " + rowData.FebPsc + "/ Mar - " + rowData.MarPsc);
                 }*/
                #endregion

                #region USE HTML AGILITY PACK

                var tableStructure = firefox2.FindElement(By.XPath(".//*[@id='plan']/tbody")).GetAttribute("innerHTML");
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(tableStructure);
                var tbRows = htmlDoc.DocumentNode.Descendants("tr");

                foreach (var tr in tbRows)
                {
                    if (tr.InnerText.Contains("Препарат")) continue;
                    if (tr.InnerText.Contains("Итого")) break;

                    var childList = tr.ChildNodes;

                    for (int k = 1; k <= month; k++)
                    {
                        var prepName = childList[1].InnerText;
                        var planTableRow = new PlanTableRow()
                        {
                            UserId = idUser,
                            PreparationName = prepName.Contains("ЛЬГОТА") ? prepName.Substring(0, prepName.Length - 9) : prepName,
                            TerritoriaName = terrName.Contains("ЛЛО") ? terrName.Substring(0, terrName.Length - 4) : terrName.Substring(0, terrName.Length - 10),
                            Month = k,
                            Pcs = Convert.ToInt32(Regex.Replace(childList[k + 2].InnerText, @"\s+", ""))
                        };
                        planForOrdinaryUser.Add(planTableRow);
                    }
                }
                #endregion

            } // закончили считывать таблицы планов для территорий
            firefox2.Navigate().GoToUrl(logout);
            return planForOrdinaryUser;
        }

        public void ReadPlanFor33BU(string url, int month, FirefoxDriver firefox2, string logout)
        {
            var arr = new List<int> { 93, 147, 1139, 1415, 1419, 142, 8015 };

            foreach (int user in arr)
            {
                var planForUser = ReadPlanTableOrdinaryUser(url, user, month, firefox2, logout);
                _planForLgotaBu33.AddRange(planForUser);
            }
        }

        public void CheckCalculatedPlans(string url, int month, FirefoxDriver firefox2, string logout)
        {
            var wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var rowList = _firefox.FindElements(By.XPath(PageElements.TableOdobrenieRowsXpath));
            var index = 0;
            for (int i = 0; i < rowList.Count; i++)
            {
                if (rowList[i].Text != "Рассылка") continue;
                index = i;
                break;
            }

            //**********************Цикл перебора БЮ*********************
            for (int j = index + 2; j <= rowList.Count; j++)
            /*for (int j = index + 4; j <= rowList.Count; j++)*/ // начинаем с БЮ 42
           //for (int j = index + 11; j <= rowList.Count; j++) // начинаем с БЮ 116
            {
                var startTime = DateTime.Now;
                var rassylkaButtonXpath = $".//*[@id='dep_info']/tbody/tr[{j}]/td[4]/input";
                var buNumder = Convert.ToInt32(_firefox.FindElement(By.XPath($".//*[@id='dep_info']/tbody/tr[{j}]/td[1]")).Text);

                Console.WriteLine("**********************");
                Console.WriteLine("ПРОВЕРКА БИЗНЕС-ЮНИТА " + buNumder);

                Helper.TryToClickWithoutException(rassylkaButtonXpath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='closeBU']")));//Close Button

                var userList = _firefox.FindElements(By.XPath(PageElements.UserTableRowsXpath));

                // ************Цикл перебора пользователей внутри БЮ**************
                #region

                for (int i = 0; i < userList.Count; i++)
                {
                    var buId = Convert.ToInt32(userList[i].GetAttribute("bunit_id"));
                    var userId = Convert.ToInt32(userList[i].GetAttribute("user_id"));
                    var userName =
                        _firefox.FindElement(By.XPath($".//*[@id='send-users-list']/tbody/tr[{i + 1}]/td[3]")).Text;
                    var regionsIdString = userList[i].GetAttribute("regions_ids");
                    var regionsIdList = new List<int>();

                    Console.WriteLine((i + 1) + ". " + userId + " " + userName + "(BU " + buId + ")");

                    try
                    {
                        if (!_spravochnikTerritorii.IsUserExistInSpravochink(userId))
                        {
                            Console.WriteLine("  " + userId + " " + userName + "(BU " + buId + ")" +
                                              " - отсутсвует в справочнике");
                            continue;
                        }
                        if (!_spravochnikTerritorii.IsBuUserSpravochikMatchPlanirovschik(buId, userId))
                        {
                            Console.WriteLine("   " + userId + " " + userName + "(BU " + buId + ")" +
                                              " - БЮ не совпадает.");
                        }

                        if (regionsIdString.Contains(','))
                        {
                            var regionsIdArray = regionsIdString.Split(',');
                            regionsIdList.AddRange(regionsIdArray.Select(t => Convert.ToInt32(t)));
                        }


                        if (!regionsIdString.Contains(','))
                        {
                            regionsIdList.Add(Convert.ToInt32(regionsIdString));
                        }

                        var regionsName =
                            _firefox.FindElement(By.XPath($".//*[@id='send-users-list']/tbody/tr[{i + 1}]/td[4]")).Text;
                        var email =
                            _firefox.FindElement(By.XPath($".//*[@id='send-users-list']/tbody/tr[{i + 1}]/td[5]")).Text;
                        var status =
                            _firefox.FindElement(By.XPath($".//*[@id='send-users-list']/tbody/tr[{i + 1}]/td[6]")).Text;

                        var user = new User()
                        {
                            BuId = buId,
                            UserId = userId,
                            UserName = userName,
                            TerritoryIdArray = regionsIdList,
                            Email = email,
                            Status = status
                        };
                        _usersListPlanirovschik.Add(user);

                        // **************CHECK EMAIL********************

                        #region

                        var userEmailSpravochnik = UserList.GetUserEmailById(user.UserId, _usersListForEmailSpravochnik);
                        if (user.Email != userEmailSpravochnik)
                        {
                            Console.WriteLine("   " + userId + " " + userName + "(BU " + buId + ") - " +
                                              "электронная почта не совпадает: (планировщик) " + user.Email + " / " +
                                              userEmailSpravochnik + " (справочник)");
                        }

                        #endregion

                        //******************CHECK TERRITORII*************************

                        #region

                        var raionTerrForUserFromSpravochnik = _spravochnikTerritorii.GetRaionTerritorr(user.UserId);
                        var oblastTerrForUserFromSpravochnik = _spravochnikTerritorii.GetOblastTerritorr(user.UserId);
                        var regionTerrForUserSpravochnik = _spravochnikTerritorii.GetRegionTerritorr(user.UserId);

                        if (buNumder == 67)
                        {
                            regionTerrForUserSpravochnik = _spravochnikTerritorii.GetRegionTerritorr(user.UserId,
                                buNumder);
                        }

                        user.TerritoryIdArray.Sort();

                        if (raionTerrForUserFromSpravochnik.Count != 0)
                        {
                            if (oblastTerrForUserFromSpravochnik.Count == 1) // если у пользователя только одна область
                            {
                                raionTerrForUserFromSpravochnik.Sort();
                                Helper.CompareIdLists(raionTerrForUserFromSpravochnik, user.TerritoryIdArray);
                            }
                            if (oblastTerrForUserFromSpravochnik.Count > 1)
                            // если у пользователя несколько областей с районами или без районов
                            {
                                foreach (var oblastId in oblastTerrForUserFromSpravochnik)
                                {
                                    var raionListForOblast =
                                        _spravochnikTerritorii.GetRaionTerritorrByOblastIdAndUserId(oblastId,
                                            user.UserId);

                                    if (raionListForOblast.Count == 0)
                                    // если у пользователя есть и области с районами и без районов
                                    {
                                        raionTerrForUserFromSpravochnik.Add(oblastId);
                                    }
                                }

                                raionTerrForUserFromSpravochnik.Sort();
                                Helper.CompareIdLists(raionTerrForUserFromSpravochnik, user.TerritoryIdArray);
                            }

                            // метод сверки районов
                        }
                        if (raionTerrForUserFromSpravochnik.Count == 0 && oblastTerrForUserFromSpravochnik.Count != 0)
                        {
                            Helper.CompareIdLists(oblastTerrForUserFromSpravochnik, user.TerritoryIdArray);
                            //метод сверки областей
                        }
                        if (raionTerrForUserFromSpravochnik.Count == 0 && oblastTerrForUserFromSpravochnik.Count == 0)
                        {
                            var oblListForUser = new List<int>();

                            foreach (var regionId in regionTerrForUserSpravochnik)
                            {
                                var tempList = _spravochnikTerritorii.GetOblastterritorrByRegionId(regionId);
                                oblListForUser.AddRange(tempList);
                            }

                            oblListForUser.Sort();
                            Helper.CompareIdLists(oblListForUser, user.TerritoryIdArray);
                            //метод сверки областей по РЕГИОНУ
                        }

                        #endregion

                        //*******************CHECK PLANS********************

                        if (!(user.Status == "Рассчитан" || user.Status == "sent"))
                        {
                            Console.WriteLine("   " + userId + " " + userName + "(BU " + buId +
                                              ") - План не расчитан (текущий статус - " + user.Status + ")");
                            continue;
                        }

                        //Console.WriteLine("   Проверка ПЛАНА:");

                        var loginCurrentUser = ReadPlanTableOrdinaryUser(url, user.UserId, month, firefox2, logout);

                        if (loginCurrentUser.Count == 0)
                        {
                            //Helper.TryToClickWithoutException($".//*[@id='closeUserBig_{userId}']", _firefox);
                            // Кнопка ОК - выход из плана пользователя
                            Console.WriteLine("   " + userId + " " + userName + "(BU " + buId +
                                              ") - отсутствует план в Планировщике");
                            continue;
                        }


                        var prosmotrPlanaButtonXpath = $".//*[@id='send-users-list']/tbody/tr[{i + 1}]/td[7]/a[1]";
                        // планы могут быть расчитаны частично и кнопка отсутсвует у нерасчитанных пользователей
                        Helper.TryToClickWithoutException(prosmotrPlanaButtonXpath, _firefox);
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath($".//*[@id='closeUserBig_{userId}']")));// Кнопка ОК в таблице пользователя
                        Thread.Sleep(500);

                        var login1340 = ReadPlanTable1340(user.UserId, month);

                        var diffName = PlanTableRowList.ComparePreparationName(login1340, loginCurrentUser);
                        // Сравнение наименований препаратов
                        if (diffName.Count > 0)
                        {
                            Console.WriteLine("Не совпадают наименования препаратов - " + diffName.Count +
                                              "кол-во несовпадений:");
                            foreach (var name in diffName)
                            {
                                Console.WriteLine(name);
                            }
                        }

                        if (!PlanTableRowList.IsTotalMatch(login1340, loginCurrentUser))
                        {
                            var error = PlanTableRowList.ComparePlans(login1340, loginCurrentUser, _planForLgotaBu33);
                            if (error.Count > 0)
                            {
                                Console.WriteLine("В плане ошибки. См. файл.");
                                File.WriteAllLines(@"D:\Sneghka\Selenium\Projects\Planirovschik\Plans errors\User_" + user.UserId + ".txt", error);
                            }
                            
                        }

                        Helper.TryToClickWithoutException($".//*[@id='closeUserBig_{userId}']", _firefox);
                        // Кнопка ОК - выход из плана пользователя
                    } //end TRY BLOCK
                    catch (Exception e)
                    {
                        Console.WriteLine("   Exception(ПОЛЬЗОВАТЕЛЬ НЕ ОБРАБОТАН) : " + userId + " " + userName + "(BU " + buId +
                                              ") - error message " + e.Message);
                        _firefox.Navigate().GoToUrl(url);
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ClosePreparationListButtonXpath)));
                        Helper.TryToClickWithoutException(PageElements.ClosePreparationListButtonXpath, _firefox);
                        Waiting.WaitForAjax(_firefox);
                        Helper.TryToClickWithoutException(PageElements.TopMenuOdobreniePlanovButtonXpath, _firefox);
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TableOdobrenieXpath)));
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='dep_info']/tbody/tr[1]/td")));
                        Helper.TryToClickWithoutException(rassylkaButtonXpath, _firefox);
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='closeBU']")));//Close Button
                        userList = _firefox.FindElements(By.XPath(PageElements.UserTableRowsXpath));
                       
                    }

                } // конец цикла перебора пользователей внутри БЮ, Проверка списка пользователей БЮ (сверка со справочником из закладки Зона ответсвенности)
                  // в планировщике могут быть пользователи, которых нет в справочнике (не ошибка)

                #endregion

                Thread.Sleep(2000);

                Helper.TryToClickWithoutException(".//*[@id='closeBU']", _firefox);//Close Button выход из списка пользователя БЮ
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TableOdobrenieXpath)));
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='dep_info']/tbody/tr[1]/td")));

                var finishTime = DateTime.Now;
                var lasting = finishTime - startTime;
                Console.WriteLine("Вермя проверки БЮ-" + buNumder + "  " + lasting);

            }// конец цикла перебора Бизнес Юнитов
        }

        #endregion
    }
}


