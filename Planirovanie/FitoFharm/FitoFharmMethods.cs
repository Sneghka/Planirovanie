using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Firefox;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Data;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace Planirovanie
{
    public class FitoFharmMethods
    {
        private readonly FirefoxDriver _firefox;
        private int numberTableRows;

        private RowDataList preparationDataSpravochnik = new RowDataList();
        private RowDataList preparationNamePlanirovschik = new RowDataList();


        public FitoFharmMethods(FirefoxDriver firefox)
        {
            _firefox = firefox;
        }


        public void StoreExcelDataAny(string path) //@"D:\Sneghka\Selenium\Projects\Planirovschik\FitoPharm.xlsx"
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path,"Select * from [Sheet1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var name = row["Полное наименование"].ToString().Trim().ToLower();
                /* var name = row["Name"].ToString().Trim().Replace("\u00A0", " ").ToLower();*/
                var rowData = new RowData
                {
                    Name = Regex.Replace(name, @"\s+", " "),
                };
                preparationDataSpravochnik.Add(rowData);
            }
        }

        public void LoginPlanirovschik(string url, string login, string password)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            _firefox.Navigate().GoToUrl(url);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SubmitButtonXPath)));
            pageElements.LoginField.SendKeys(login);
            pageElements.PasswordField.SendKeys(password);
            pageElements.SubmitButton.Click();
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='dialog_init']")));
            Thread.Sleep(2000);
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
                    _firefox.FindElement(By.XPath(".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[2]"))
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
        }

        public void CompareWebWithExcel()
        {
            var difference = RowDataList.CompareStrings(preparationNamePlanirovschik.GetNamesList(),
                preparationDataSpravochnik.GetUniqueNames());
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

        public void CompareExcelWithWeb()
        {
            var difference = RowDataList.CompareStrings(preparationDataSpravochnik.GetUniqueNames(),
                preparationNamePlanirovschik.GetNamesList());
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
                var raschetButtonXPath = ".//*[@id='preparation_info']/tbody/tr[" + i + "]/td[5]/input[1]";
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
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(PageElements.SavePlanButtonXPath)));
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

        public void LogoutFitoFharm(string url)
        {
            _firefox.Navigate().GoToUrl(url);
        }

    }
}
