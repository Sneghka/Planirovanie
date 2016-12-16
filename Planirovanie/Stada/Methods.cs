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

        private int numberTableRows;
        private RowDataList preparationNamePlanirovschik = new RowDataList();
        private RowDataList preparationDataSpravochnik = new RowDataList();
        private RowTerritoriiList planirovschikTerritorii = new RowTerritoriiList();
        private RowTerritoriiList spravochnikTerritorii = new RowTerritoriiList();
        private RowTerritoriiList planirovschikBuId = new RowTerritoriiList();
        private RowTerritoriiList spravochnikBuId = new RowTerritoriiList();
        private List<RowTerritorii> differencePlanirovschikWithSpravochik = new RowTerritoriiList();
        private List<RowTerritorii> differenceSpravochikWithPlanirovschik = new RowTerritoriiList();
        private RowDataList _distribution2015XlsList = new RowDataList();
        private RowDataList _distribution2016XlsList = new RowDataList();
        private RowDataList _audit2015XlsList = new RowDataList();
        private List<string> handles;
        private string planirovschikdWindow;
        private string dashBoardWindow;
        private List<LoginPassword> loginPasswordList = new List<LoginPassword>();
        private List<string> grListValue = new List<string>();


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
                preparationDataSpravochnik.Add(rowData);
            }
        }

        public void StoreExcelDataBuTerritorii(string path) //@"D:\Sneghka\Selenium\Projects\Planirovschik\FitoPharm.xlsx"
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path,
                "Select * from [zone_of_resp$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
              
                var rowData = new RowData
                {
                    Id_BU = Convert.ToInt32(row["BUID"]),
                    DistrictName3 = row["District.Name3"].ToString(),
                    IdSotr = Convert.ToInt32(row["id_Sotr"]),
                    FIO = row["Full_name"].ToString(),
                    Position = row["Position"].ToString()
                    };
                preparationDataSpravochnik.Add(rowData);
            }
        }

       public void StoreExcelDataAny(string path)//@"D:\Sneghka\Selenium\Projects\Planirovschik\FitoPharm.xlsx"
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path,
                "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var name = row["Полное наименование"].ToString().Trim().ToLower();
                /* var name = row["Name"].ToString().Trim().Replace("\u00A0", " ").ToLower();*/
                var rowData = new RowData
                {
                    Name = Regex.Replace(name, @"\s+", " "),
                    /* Name = Regex.Replace(name, @"\s+", " "),
                     Id_BU = Convert.ToInt32(row["id_BU"]),*/

                };
                preparationDataSpravochnik.Add(rowData);
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
          /* wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='dialog_init']")));*/
            Thread.Sleep(5000);
        }

        public bool IsLoginSuccess(string url, string login, string password)
        {
            if (_firefox.FindElement(By.XPath(".//*[@id='dialog-confirm']")).GetAttribute("style") == "display: none;")
                return true;
            return false;
        }

        public bool IsPreparationListExist()
        {
            if (Helper.IsElementPresent(By.XPath("html/body/div[4]/div[3]/div/button[1]"), _firefox)) //кнопка "Закрыть" на списке препаратов
                return true;
            return false;
        }

        public void StorePreparationNamesFromPlanirovschik()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(1000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            numberTableRows = tableRows.Count;
            Thread.Sleep(4000);
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице Планировщика");
            for (int i = 1; i <= numberTableRows; i++)
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
                    Status = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[3]")).GetAttribute("aria-disabled")

            };
                preparationNamePlanirovschik.Add(rowData);
            }
        }

        public void GetListPreparationFromExcel(int[] months)
        {
            Debug.WriteLine("Список препаратов из Екселя:");
            foreach (var name in preparationDataSpravochnik.GetUniqueWebNames(months))
            {
                Debug.WriteLine(name);
            }
        }

       public void ComparePreparationNameThroughObjects(int[] months)
        {
            var convertSpravochnik = RowDataList.ConvertSpravochikList(months, preparationDataSpravochnik);
            Console.WriteLine("СПРАВОЧНИК");
         
           var diff1 = RowDataList.CompareRowDataObjects(convertSpravochnik, preparationNamePlanirovschik);
            if (diff1.Count != 0)
            {
                Console.WriteLine("Данные из справочника отсутствуют в планировщике:");
                foreach (var d in diff1)
                {
                    Console.WriteLine(d.IdPrUniq + " " + d.Name + " (BU_ID - " + d.Id_BU + "; Segment - "  + d.Segment + "; Group - " + d.Group + ")");
                }
            }
            else
            {
                Console.WriteLine("Сверка справочника с планировщиком. Расхождений нет");
            }
            var diff2 = RowDataList.CompareRowDataObjects(preparationNamePlanirovschik, convertSpravochnik);
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
            var convertSpravochnik = RowDataList.ConvertSpravochikList(months, preparationDataSpravochnik);
            var convertSpravochnikWithAutoplanOnly = RowDataList.GetPreparationWithAutoPlanFromSpravochnik(convertSpravochnik);
            var planirovschikWithAutoplanOlny = RowDataList.GetPreparationWithAutoPlanFromPlanirovschik(preparationNamePlanirovschik);

          
            var diff1 = RowDataList.CompareRowDataObjects(convertSpravochnikWithAutoplanOnly, planirovschikWithAutoplanOlny);
            if (diff1.Count != 0)
            {
                Console.WriteLine("Данные из справочника отсутствуют в планировщике:");
                foreach (var d in diff1)
                {
                    Console.WriteLine(d.IdPrUniq + " " + d.Name + " (BU_ID - " + d.Id_BU + "; Segment - " + d.Segment + "; Group - " + d.Group + ")");
                }
            }
            else
            {
                Console.WriteLine("Сверка справочника с планировщиком. Расхождений нет");
            }
            var diff2 = RowDataList.CompareRowDataObjects(planirovschikWithAutoplanOlny, convertSpravochnikWithAutoplanOnly);
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
            var difference = RowDataList.CompareStrings(preparationNamePlanirovschik.GetNamesList(),preparationDataSpravochnik.GetUniqueWebNames(months));
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
            var difference = RowDataList.CompareStrings(preparationDataSpravochnik.GetUniqueWebNames(months),preparationNamePlanirovschik.GetNamesList());
         

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
                Console.WriteLine(month + "month pcs: справочник" + pcsSpravochnik + " = " + pcsPlanirovschik + " планировщик");
            }
            else
            {
                Console.WriteLine(month + "month pcs: справочник" + pcsSpravochnik + " НЕ РАВНО!!! " + pcsPlanirovschik + " планировщик");
            }
        }

        public void MessageCheckPreparationMethodTotal(int preparationId, int[] months, decimal totalSum, int totalPcs)
        {
            if ((preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) - totalSum) < 5 && preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) - totalSum > -5)
            {
                Console.WriteLine("Total sum: справочник " + preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) + " = " + totalSum + " планировщик");
            }
            else
            {
                Console.WriteLine("Total sum: справочник " + preparationDataSpravochnik.GetTotalSumRubById(preparationId, months) + " НЕ РАВНО !!!! " + totalSum + " планировщик");
            }

            if (preparationDataSpravochnik.GetTotalSumPcsById(preparationId, months) == totalPcs)
            {
                Console.WriteLine("Total pcs: справочник " + preparationDataSpravochnik.GetTotalSumPcsById(preparationId, months) + " = " + totalPcs + " планировщик");
            }
            else
            {
                Console.WriteLine("Total pcs: справочник " + preparationDataSpravochnik.GetTotalSumPcsById(preparationId, months) + " НЕ РАВНО!!! " + totalPcs + " планировщик");
            }
        }

        public void CheckPreparationData(int[] months)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));// get list of preparation
            numberTableRows = tableRows.Count;
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= numberTableRows; i++)
            {
                Console.WriteLine("№" + i);
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationBuId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("bu_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
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
                        var totalPcsMonthLgota = Convert.ToInt32(pageElements.GetPcsLgotaMonth(month).Text.Replace(" ", ""));
                        monthSumLgota.Add(month, totalPcsMonthLgota);
                    }

                    foreach (KeyValuePair<int, int> kvp in monthSumLgota)
                    {
                        MessageCheckPreparationMethodByMonth(kvp.Key, preparationDataSpravochnik.GetPcsByIdAndSegmentAndMonth(preparationId, kvp.Key, 2), kvp.Value);
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
                        totalSum = Convert.ToDecimal(pageElements.TotalSumRub.Text.Substring(0, pageElements.TotalSumRub.Text.Length - 5).Replace(" ", "").Replace(".", ","));
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
                        MessageCheckPreparationMethodByMonth(kvp.Key, preparationDataSpravochnik.GetSumPcsByIdAndMonth(preparationId, kvp.Key), kvp.Value);
                    }
                }
               
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                Thread.Sleep(500);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.FindPreparationInputFieldXPath)));

            } //конец цикла FOR перебора всех препаратов

        } //конец метода

        public void CheckPreparationListForPM(int user)
        {
            preparationNamePlanirovschik.Clear();

            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));// get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);

            for (int i = 1; i <= numberTableRows; i++)
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
                preparationNamePlanirovschik.Add(rowData);
            }

            var listPreparationIDSpravochnik = preparationDataSpravochnik.GetIdListByUserWithoutAutoplan(user);
            var listPreparationIDPlanirovschik = preparationNamePlanirovschik.GetIdList();

            var compareWebwithExcel = RowDataList.CompareNumbers(listPreparationIDPlanirovschik, listPreparationIDSpravochnik);
            var compareExcelWithWeb = RowDataList.CompareNumbers(listPreparationIDSpravochnik, listPreparationIDPlanirovschik);

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
            var listPreparationIDSpravochnik = preparationDataSpravochnik.GetIdListByUserWithoutAutoplan(user);

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
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\Sneghka\Instructions NEW\Planirovschik\Current_users_territory_1.xls", "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["FIO"].ToString(),
                    IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                    Position = row["Position"].ToString(),
                    DistrictName = row["DistrictName"].ToString()

                };
                planirovschikTerritorii.Add(rowData);
            }
        }

        public void StoreExcelDataTerritoriiSpravochnik()
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\Sneghka\Instructions NEW\Planirovschik\Spravochnik_terr.xlsx", "Select * from [zone_of_resp$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["Full_name"].ToString(),
                    IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                    Position = row["Position"].ToString(),
                    DistrictName = row["DistrictName2"].ToString()

                };
                spravochnikTerritorii.Add(rowData);
            }
        }

        public void CompareTerritoriiSpravochnikWithPlanirovschik()
        {
            differenceSpravochikWithPlanirovschik = RowTerritoriiList.CompareTerritoriis(spravochnikTerritorii, planirovschikTerritorii);
            Console.WriteLine("Данные есть в Справочнике, но отсутствуют в планировщике");
            /* foreach (var x in differenceSpravochikWithPlanirovschik)
                Console.WriteLine(x.Position + "/ "+ x.FIO + " /" + x.DistrictName);*/
        }

        public void CompareTerritoriiPlanirovschikWithSpravochnik()
        {
            differencePlanirovschikWithSpravochik = RowTerritoriiList.CompareTerritoriis(planirovschikTerritorii, spravochnikTerritorii);
            var y = differencePlanirovschikWithSpravochik.Count;
            Console.WriteLine("Данные есть в Планировщике, но отсутствуют в Справочнике");
            /* foreach (var z in differencePlanirovschikWithSpravochik)
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

            for (int i = 3; i <= differencePlanirovschikWithSpravochik.Count; i++)
            {

                ws.Cells[i, 1] = differencePlanirovschikWithSpravochik[i - 3].Position;
                ws.Cells[i, 2] = differencePlanirovschikWithSpravochik[i - 3].FIO;
                ws.Cells[i, 3] = differencePlanirovschikWithSpravochik[i - 3].DistrictName;
            }

            ws.Cells[1, 5] = "Есть в справочнике, но нет в планировщике";
            ws.Cells[2, 5] = "Position";
            ws.Cells[2, 6] = "Fio";
            ws.Cells[2, 7] = "DistrictName";

            for (int i = 3; i <= differenceSpravochikWithPlanirovschik.Count; i++)
            {

                ws.Cells[i, 5] = differenceSpravochikWithPlanirovschik[i - 3].Position;
                ws.Cells[i, 6] = differenceSpravochikWithPlanirovschik[i - 3].FIO;
                ws.Cells[i, 7] = differenceSpravochikWithPlanirovschik[i - 3].DistrictName;
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
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\Sneghka\Instructions NEW\Planirovschik\Current_preparation_bu.xls", "Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["FIO"].ToString(),
                    IdSotr = Convert.ToInt32(row["id"]),
                    BuId = row["roleName"].ToString()

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
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\Sneghka\Instructions NEW\Planirovschik\Spravochnik_bu.xlsx", "Select * from [zone_of_resp$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new RowTerritorii()
                {
                    FIO = row["Full_name"].ToString(),
                    IdSotr = Convert.ToInt32(row["Id_Sotr"]),
                    BuId = row["BUID"].ToString()

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
                if (!(z.BuId == "81" || z.BuId == "82"))
                {
                    newDifference.Add(z);
                }
                else
                {
                    foreach (var c in difference)
                    {
                        if (c.IdSotr == z.IdSotr && c.BuId == "114") newDifference.Add(z);
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

            WorkWithExcelFile.ExcelFileToDataTable(out dt2015, @"D:\Sneghka\Selenium\Projects\Planirovschik\Disrtibution_total.xlsx",
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

            WorkWithExcelFile.ExcelFileToDataTable(out dt2016, @"D:\Sneghka\Selenium\Projects\Planirovschik\Disrtibution_total.xlsx",
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

        public void CheckDistributionDataWithExcel()
        {
            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(1000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));// get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Console.WriteLine(numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= numberTableRows; i++)
            {

                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButton = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
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

                // Блок сбора данных за 2015 год

                /* wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2015Xpath)));
                 TryToClickWithoutException(PageElements.SalesData2015Xpath, pageElements.SalesData2015);
                 WaitPatternPresentInAttribute(PageElements.SalesData2015Xpath, "class", "ui-tabs-selected");
                 Thread.Sleep(200);
                 var total2015 = Convert.ToInt32(pageElements.TotalSumSpravochyeDannye2015.Text.Replace(" ", ""));
                 if (total2015 == _distribution2015XlsList.GetUpakovkiById(preparationId))
                 {
                     Console.WriteLine(preparationName + "_2015 (web/xls): " + total2015 + " = " +
                                       _distribution2015XlsList.GetUpakovkiById(preparationId));
                 }
                 else
                 {
                     Console.WriteLine(preparationName + "_2015 (web/xls): " + total2015 + " НЕ РАВНО!!!! " +
                                       _distribution2015XlsList.GetUpakovkiById(preparationId));
                 }
 */
                // Блок сбора данных за 2016 год

                Helper.TryToClickWithoutException(PageElements.SalesData2016Xpath, _firefox);
                Waiting.WaitPatternPresentInAttribute(PageElements.SalesData2016Xpath, "class", "ui-tabs-selected", _firefox);
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
            planirovschikdWindow = _firefox.CurrentWindowHandle;
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
            _firefox.Navigate().GoToUrl("http://pharmxplorer.com.ua/QvAJAXZfc/opendoc.htm?document=TestDocs/stada/xls_data_test.qvw&host=QVS@qlikview&anonymous=true");
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
            dashBoardWindow = _firefox.CurrentWindowHandle;
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


            WorkWithExcelFile.ExcelFileToDataTable(out dtAudit2015, @"D:\Sneghka\Instructions NEW\Planirovschik\Audit.xlsx",
                "Select * from [Audit2015_1$]");
            foreach (DataRow row in dtAudit2015.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var rowData = new RowData
                {
                    IdPrUniq = Convert.ToInt32(row["preparationId"]),
                    Name = row["name"].ToString(),
                    Upakovki = Convert.ToInt32(row["Свои упаковки"]),
                    UpakovkiConcurent = Convert.ToInt32(row["Конкурентные упаковки"])
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
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= numberTableRows; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[2]")).Text;


                if (!_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]")).GetAttribute("class").Contains("ui-button-disabled")) // начало проверки на активность кнопки РАСЧЁТ  
                {
                    ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]")));
                    Thread.Sleep(500);
                    Helper.TryToClickWithoutException(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]", _firefox);
                    // click "Расчёт" для выбранного элемента
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AuditDataOwn2015XPath)));
                    Helper.TryToClickWithoutException(PageElements.AuditDataOwn2015XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataOwn2015XPath, "class", "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);
                    var totalOwn2015 = Convert.ToInt32(pageElements.TotalSumOwnSalesData2015.Text.Replace(" ", ""));

                    if (totalOwn2015 == _audit2015XlsList.GetUpakovkiById(preparationId))
                    {
                        Console.WriteLine(preparationName + "_2015Own (web/xls): " + totalOwn2015 + " = " +
                                          _audit2015XlsList.GetUpakovkiById(preparationId));
                    }
                    else
                    {
                        Console.WriteLine(preparationName + "_2015Own (web/xls): " + totalOwn2015 + " НЕ РАВНО!!! " +
                                         _audit2015XlsList.GetUpakovkiById(preparationId));
                    }

                    Helper.TryToClickWithoutException(PageElements.AuditDataCompetitor2015XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataCompetitor2015XPath, "class", "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);
                    var totalCompetitior2015 = Convert.ToInt32(pageElements.TotalSumCompetitorSalesData2015.Text.Replace(" ", ""));

                    if (totalCompetitior2015 == _audit2015XlsList.GetUpakovkiConcurentById(preparationId))
                    {
                        Console.WriteLine(preparationName + "_2015Competitor (web/xls): " + totalCompetitior2015 + " = " +
                                          _distribution2016XlsList.GetUpakovkiConcurentById(preparationId));
                    }
                    else
                    {
                        Console.WriteLine(preparationName + "_2015Competitor (web/xls): " + totalCompetitior2015 + " НЕ РАВНО!!!! " +
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

        public void CheckAuditDataWithDashBoard()
        {
            var action = new Actions(_firefox);
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);

            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));
            // get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= numberTableRows; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                Console.WriteLine("№" + i);
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[2]")).Text;

                if (!_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]")).GetAttribute("class").Contains("ui-button-disabled")) // начало проверки на активность кнопки РАСЧЁТ  
                {

                    ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]")));
                    Thread.Sleep(500);
                    Helper.TryToClickWithoutException(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]", _firefox);
                    // click "Расчёт" для выбранного элемента
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SpravochyeDannyeButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.SpravochyeDannyeButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SalesData2015Xpath)));
                    Helper.TryToClickWithoutException(PageElements.AuditDataOwn2015XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataOwn2015XPath, "class", "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);

                    var totalOwn2015Plan = Convert.ToInt32(pageElements.TotalSumOwnSalesData2015.Text.Replace(" ", ""));

                    Helper.TryToClickWithoutException(PageElements.AuditDataCompetitor2015XPath, _firefox);
                    Waiting.WaitPatternPresentInAttribute(PageElements.AuditDataCompetitor2015XPath, "class", "ui-tabs-selected", _firefox);
                    Thread.Sleep(200);

                    var totalCompetitior2015Plan = Convert.ToInt32(pageElements.TotalSumCompetitorSalesData2015.Text.Replace(" ", ""));

                    Thread.Sleep(2000);

                    _firefox.SwitchTo().Window(dashBoardWindow);

                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchPreperationIdAuditWebXPath)));

                    Helper.TryToClickWithoutException(PageElements.SearchPreperationIdAuditWebXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.InputFieldAuditXPath)));
                    pageElements.InputFieldAuditWeb.SendKeys("(" + preparationId + ")" + Keys.Enter);
                    Waiting.WaitPatternPresentInAttribute(".//*[@id='57']/div[3]/div/div[1]/div[1]", "class", "QvSelected", _firefox);
                    Thread.Sleep(2000);
                    var totalOwn2015Dash = Convert.ToInt32(pageElements.TotalOwnPcsAuditWeb.GetAttribute("title"));
                    var totalCompetitor2015Dash = Convert.ToInt32(pageElements.TotalCompetitorPcsAuditWeb.GetAttribute("title"));

                    if (totalOwn2015Plan == totalOwn2015Dash)
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " + totalOwn2015Plan + " = " + totalOwn2015Dash);
                    }
                    else
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " + totalOwn2015Plan + " НЕ РАВНО!!! " + totalOwn2015Dash);
                    }


                    if (totalCompetitior2015Plan == totalCompetitor2015Dash)
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " + totalCompetitior2015Plan + " = " + totalCompetitor2015Dash);
                    }
                    else
                    {
                        Console.WriteLine(preparationName + "_2015 (planirovschik/dashboard): " + totalCompetitior2015Plan + " НЕ РАВНО!!!! " + totalCompetitor2015Dash);
                    }
                    _firefox.SwitchTo().Window(planirovschikdWindow);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.RaschetPlanaButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.RaschetPlanaButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                }// end IF check Расчет button
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


            WorkWithExcelFile.ExcelFileToDataTable(out dtLoginPassword, @"D:\Sneghka\Instructions NEW\Planirovschik\Check_Login_Pass.xlsx",
                "Select * from [Пользователи$]");
            foreach (DataRow row in dtLoginPassword.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var loginPassword = new LoginPassword()
                {
                    Login = row["login"].ToString(),
                    Password = row["parole"].ToString()
                };
                loginPasswordList.Add(loginPassword);
            }
            Console.WriteLine("Login/Password was stored");
        }

        public void CheckLoginPasswordMethod1(string url)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            _firefox.Navigate().GoToUrl(url);
            int i = 1;
            for (int a = 189; a < loginPasswordList.Count; a++)
            {
               
                pageElements.LoginField.Clear();
                pageElements.PasswordField.Clear();
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
                pageElements.LoginField.SendKeys(loginPasswordList[a].Login);
                pageElements.PasswordField.SendKeys(loginPasswordList[a].Password);
                pageElements.SubmitButton.Click();
               
                Thread.Sleep(4000);
                if (_firefox.FindElement(By.XPath(".//*[@id='dialog-confirm']")).GetAttribute("style") == "display: none;")
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

                    Console.WriteLine("№" + i + "  Ok: " + loginPasswordList[a].Login + " / " + loginPasswordList[a].Password);
                    i++;
                }

                else
                {
                    Console.WriteLine("№" + i + "  Incorrect login or password: " + loginPasswordList[a].Login + " / " + loginPasswordList[a].Password);
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
            for (int a = 0; a < loginPasswordList.Count; a++)
            {
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
                pageElements.LoginField.SendKeys(loginPasswordList[a].Login);
                pageElements.PasswordField.SendKeys(loginPasswordList[a].Password);
                pageElements.SubmitButton.Click();
                //Waiting.WaitForAjax()();
                Waiting.WaitForAjax(_firefox);
                if (_firefox.FindElement(By.XPath(".//*[@id='dialog-confirm']")).GetAttribute("style") == "display: none;")
                {
                    Console.WriteLine("№" + i + "  Ok: " + loginPasswordList[a].Login + " / " + loginPasswordList[a].Password);
                    i++;
                }
                else
                {
                    Console.WriteLine("№" + i + "  Incorrect login or password: " + loginPasswordList[a].Login + " / " + loginPasswordList[a].Password);
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
            Thread.Sleep(1000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr")); // get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= numberTableRows; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[2]")).Text;
                var preparationBu = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("bu_id"));
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
                    Console.WriteLine(preparationId + " " + preparationName + " - статус - " + raschetButtonPlanStatus);
                    continue;
                }

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", _firefox.FindElement(By.XPath(raschetButtonXPath)));
                Thread.Sleep(1500);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(raschetButtonXPath)));
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);

                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SavePlanButtonXPath)));
                if (_firefox.FindElement(By.XPath(".//*[@id='save_plan_customer']")).GetAttribute("aria-disabled") ==
                    "true")
                {
                    Console.WriteLine("№" + i + " " + preparationName + " - препарат НЕ утверждён. Кнопка 'Сохранить план' неактивна");
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                    Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                    wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='dialog_init']")));
                    continue;
                }

                Helper.TryToClickWithoutException(PageElements.SavePlanButtonXPath, _firefox);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AcceptButtonXpath)));
                Helper.TryToClickWithoutException(PageElements.AcceptButtonXpath, _firefox);
                Thread.Sleep(200);

                _firefox.FindElement(By.XPath("/html/body/div[@class='ui-pnotify ']/div/div[4]/center/input")).Click(); // click "Перейти к утверждению"
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

            _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[1]/td[6]/input[2]")).Click(); // First element's PlanButton 
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='plan_list']")));
            var tableRowsPlansApprove = _firefox.FindElements(By.XPath(".//*[@id='preparation_info_short']/tbody/tr")); // get list of preparation
            var numberTableRowsPlansApprove = tableRowsPlansApprove.Count;
            Thread.Sleep(4000);
            Debug.WriteLine(numberTableRowsPlansApprove + " кол-во строк в таблице");

            for (int i = 1; i <= numberTableRowsPlansApprove; i++)
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info_short']")));
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info_short']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationNameXPath = ".//*[@id='preparation_info_short']/tbody/tr[" + i + "]/td[2]";
                var preparationName = _firefox.FindElement(By.XPath(preparationNameXPath)).Text;
                var preparationStatusXPath = ".//*[@id='preparation_info_short']/tbody/tr[" + i + "]/td[3]";
                var preparationStatus = _firefox.FindElement(By.XPath(preparationStatusXPath));

                ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", _firefox.FindElement(By.XPath(preparationNameXPath)));
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(preparationNameXPath, _firefox);// click выбранный элемент

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
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));// get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            //Debug.WriteLine(numberTableRows + " кол-во строк в таблице");
            for (int i = 1; i <= numberTableRows; i++)
            {
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButton = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";

                if (raschetButton.GetAttribute("class").Contains("ui-button-disabled"))
                {
                    Console.WriteLine(preparationName + " - кнопка Расчет неактивна");
                    continue;
                }
               ((IJavaScriptExecutor)_firefox).ExecuteScript("arguments[0].scrollIntoView(true);", raschetButton);
                Thread.Sleep(500);
                Helper.TryToClickWithoutException(raschetButtonXPath, _firefox);// click "Расчёт" для выбранного элемента
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ChoosePreparationButtonXPath)));
                Thread.Sleep(500);
                var valueGr = _firefox.ExecuteScript("return document.getElementById('COMP_GR').previousSibling.innerHTML;"); // instead innerHTML can use innerContent
                grListValue.Add("№" + i + " " + preparationId + " " + preparationName + " -- " + valueGr);
                Console.WriteLine("№" + i + " " + preparationId + " " + preparationName + " -- " + valueGr);
                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            }
            File.WriteAllLines(@"D:\Sneghka\Selenium\Created documents\GR_Value.doc", grListValue);
        }

        public void PrintGR()
        {
            File.WriteAllLines(@"D:\Sneghka\GR_Value.doc", grListValue);
        }

        public void IsGrUnchangeable()
        {

            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            Thread.Sleep(2000);
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));// get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= numberTableRows; i++)
            {

                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                if (!preparationDataSpravochnik.GetIdList().Contains(preparationId)) continue;

                var preparationBuId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("bu_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));


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
                    Console.WriteLine("№" + i + " " + preparationId + " " + " " + preparationName + ": Слайдер НЕАКТИВЕН");
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
            var tableRows = _firefox.FindElements(By.XPath(".//*[@id='preparation_info']/tbody/tr"));// get list of preparation
            numberTableRows = tableRows.Count;
            Thread.Sleep(2000);
            Debug.WriteLine(numberTableRows + " кол-во строк в таблице");

            for (int i = 1; i <= numberTableRows; i++)
            {

                var preparationId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("data_id"));
                if (!preparationDataSpravochnik.GetIdList().Contains(preparationId)) continue;

                var preparationBuId = Convert.ToInt32(_firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]")).GetAttribute("bu_id"));
                var preparationName = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[3]")).Text;
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]";
                var raschetButton = _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[6]/input[1]"));


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
                    Console.WriteLine("№" + i + " " + preparationId + " " + " " + preparationName + ": Слайдер НЕАКТИВЕН");
                }

                Helper.TryToClickWithoutException(PageElements.ChoosePreparationButtonXPath, _firefox);
                wait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(".//*[@id='preparation_info']/tbody")));
            }

        }

        #endregion
    }
}

