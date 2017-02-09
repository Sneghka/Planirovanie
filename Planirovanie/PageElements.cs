using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace Planirovanie
{
    public class PageElements
    {
        private readonly FirefoxDriver _firefox;

        public PageElements(FirefoxDriver firefox)
        {
            _firefox = firefox;

        }

        #region ElementsXPath


        public const string LoginFieldXPath = ".//*[@id='login']";
        public const string PasswordFieldXPath = ".//*[@id='password']";
        public const string SubmitButtonXPath = ".//*[@id='center']/form/fieldset/input[3]";
        public const string PreparationTable = ".//*[@id='preparation_info']/tbody";
        public const string TotalSumRubXPath = ".//*[@id='sumEuro']";
        public const string TotalPcsXPath = ".//*[@id='sumPOPPVP']";
        public const string ChoosePreparationButtonXPath = ".//*[@id='load_sku_customer']";
        public const string FindPreparationInputFieldXPath = ".//*[@id='preparation_info_filter']/input";
        public const string SpravochyeDannyeButtonXPath = ".//*[@id='tabs']/ul/li[2]";
        public const string SalesData2016Xpath = ".//*[@id='tab_info']/ul/li[1]";
        public const string SalesData2017Xpath = ".//*[@id='tab_info']/ul/li[2]";
        public const string RaschetPlanaButtonXPath = ".//*[@id='tabs']/ul/li[1]/a";
        public const string AuditDataOwn2016XPath = ".//*[@id='tab_info']/ul/li[3]";
        public const string AuditDataCompetitor2016XPath = ".//*[@id='tab_info']/ul/li[5]";
        public const string SearchAreaNameAuditWebXPath = ".//*[@class='QvFrame Document_LB06']/div[2]/div[1]/div";
        public const string SearchPeriodAuditWebXPath = ".//*[@class='QvFrame Document_LB04']/div[2]/div[1]/div";
        public const string InputFieldAuditXPath = "html/body/div[2]/input";
        public const string GrXPath = ".//*[@class='inputShadow']";
        public const string GrSliderXPath = ".//*[@id='sliderGP']";
        public const string ClosePreparationListButtonXpath = "html/body/div[4]/div[3]/div/button[1]";
        public const string TopMenuOdobreniePlanovButtonXpath = ".//*[@id='tabs']/ul/li[6]/a";
        public const string TableOdobrenieXpath = ".//*[@id='dep_info']";
        public const string TableOdobrenieRowsXpath = ".//*[@id='dep_info']/tbody/tr";
        public const string GlobalApprovePlanButton1340Xpath = ".//*[@id='dep_info']/tbody/tr[2]/td[4]/input";
        public const string UserTableRowsXpath = ".//*[@id='send-users-list']/tbody/tr";






        public const string Gr1340XPath =".//*[@id='customer_settings_accordion']/div/table/tbody/tr[1]/td[3]/div/div[2]/div";

        public const string AreaLevel_2AuditWebXPath =".//*[@class='QvFrame Document_LB02']/div[3]/div/div[1]/div[3]/div[1]";

        public const string LockAuditWebXPath = ".//*[@id='QvAjaxToolbar']/ul[2]/li[13]/a";
        public const string SearchPreperationIdAuditWebXPath = ".//*[@class='QvFrame Document_LB03']/div[2]/div[1]/div";
        public const string SavePlanButtonXPath = ".//*[@id='save_plan_customer']";
        public const string AcceptButtonXpath = ".//*[@aria-labelledby='ui-dialog-title-dialog-plan-settings']/div[3]/div/button[2]";
        public const string GluePopupGoToconfirmationXPath = "/html/body/div[@class='ui-pnotify']/div/div[3]";

        public const string PopupGoToconfirmationButtonXPath =
            "html/body/div[@class='ui-pnotify']/div/div[4]/center/input";

        public const string AcceptPlanButtonXPath = ".//*[@id='accept_plan']";

        public const string ConfirmPlanButtonXPath =
            "html/body/div[@aria-labelledby='ui-dialog-title-dialog-confirm']/div[3]/div/button[1]";

        public const string PlansConfirmationXPath = ".//*[@id='tabs']/ul/li[3]/a";
        public const string RefreshButtonXPath = ".//*[@id='reload_plans']";
        public const string ApprovePlanButtonXPath = ".//*[@id='approve_plan']";




        #endregion


        public IWebElement LoginField
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='login']")); }
        }

        public IWebElement PasswordField
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='password']")); }
        }

        public IWebElement SubmitButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='center']/form/fieldset/input[3]")); }
        }

        public IWebElement TotalSumRub
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='sumEuro']")); }
        }

        public IWebElement TotalPcs
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='sumPOPPVP']")); }
        }


        private string[] _PcsXPath =
       {
             "",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[14]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[15]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[16]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[17]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[18]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[19]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[20]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[21]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[22]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[23]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[24]",
            ".//*[@id='tableres_customer']/tbody/tr[7]/td[25]"
        };

        public IWebElement GetPcsMonth(int n)
        {
            return _firefox.FindElement(By.XPath(_PcsXPath[n]));
        }


        public IWebElement PcsJanuary
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[12]")); }
        }

        public IWebElement PcsFebruary
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[13]")); }
        }

        public IWebElement PcsMarch
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[14]")); }
        }

        public IWebElement PcsApril
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[15]")); }
        }

        public IWebElement PcsMay
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[16]")); }
        }

        public IWebElement PcsJune
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[17]")); }
        }

        public IWebElement PcsJuly
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[18]")); }
        }

        public IWebElement PcsAugust
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[19]")); }
        }

        public IWebElement PcsSeptember
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[20]")); }
        }

        public IWebElement PcsOctober
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[21]")); }
        }

        public IWebElement PcsNovember
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[22]")); }
        }

        public IWebElement PcsDecember
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[7]/td[23]")); }
        }


        #region Months for Lgota

        private string[] _PcsLgotaXPath =
        {
            "",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[14]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[15]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[16]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[17]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[18]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[19]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[20]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[21]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[22]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[23]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[24]",
            ".//*[@id='tableres_customer']/tbody/tr[10]/td[25]"
        };

        public IWebElement GetPcsLgotaMonth(int n)
        {
            return _firefox.FindElement(By.XPath(_PcsLgotaXPath[n]));
        }


        public IWebElement PcsJanuaryLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[12]")); }
        }
        public IWebElement PcsFebruaryLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[13]")); }
        }
        public IWebElement PcsMarchLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[14]")); }
        }
        public IWebElement PcsAprilLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[15]")); }
        }
        public IWebElement PcsMayLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[16]")); }
        }
        public IWebElement PcsJuneLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[17]")); }
        }
        public IWebElement PcsJulyLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[18]")); }
        }
        public IWebElement PcsAugustLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[19]")); }
        }
        public IWebElement PcsSeptemberLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[20]")); }
        }
        public IWebElement PcsOctoberLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[21]")); }
        }
        public IWebElement PcsNovemberLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[22]")); }
        }
        public IWebElement PcsDecemberLgota
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableres_customer']/tbody/tr[10]/td[23]")); }
        }

        #endregion

        #region DashBoardAudit

        public IWebElement LockAuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='QvAjaxToolbar']/ul[2]/li[13]/a")); }
        }

        public IWebElement SearchAreaNameAuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB06']/div[2]/div[1]/div")); }
        }
        public IWebElement AreaLevel_2AuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB02']/div[3]/div/div[1]/div[3]/div[1]")); }
        }
        public IWebElement InputFieldAuditWeb
        {
            get { return _firefox.FindElement(By.XPath("html/body/div[2]/input")); }
        }
        public IWebElement SearchPeriodAuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB04']/div[2]/div[1]/div")); }
        }
        public IWebElement SearchPreperationIdAuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB03']/div[2]/div[1]/div")); }
        }
        public IWebElement TotalOwnPcsAuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_CH01']/div[3]/div[1]/div[1]/div[4]/div/div[7]")); }
        }
        public IWebElement TotalCompetitorPcsAuditWeb
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_CH01']/div[3]/div[1]/div[1]/div[4]/div/div[8]")); }
        }

        public IWebElement Gr
        {
            //get {return _firefox.FindElement(By.XPath(".//*[@id='customer_settings_accordion']/div/table/tbody/tr[1]/td[3]/div/div[2]/div"));}
            get { return _firefox.FindElement(By.XPath(".//*[@class='inputShadow']")); }
        }

        public IWebElement GrSlider
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='sliderGP']")); }
        }

        public IWebElement Gr1340
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='customer_settings_accordion']/div/table/tbody/tr[1]/td[3]/div/div[2]/div")); } // *[@id='COMP_GR'] 
        }

        #endregion

        public IWebElement RaschetPlanaButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tabs']/ul/li[1]/a")); }
        }

        public IWebElement ChoosePreparationButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='load_sku_customer']")); }
        }

        public IWebElement FindPreparationInputField
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='preparation_info_filter']/input")); }
        }

        public IWebElement SavePlanButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='save_plan_customer']")); }
        }

        public IWebElement AcceptButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@aria-labelledby='ui-dialog-title-dialog-plan-settings']/div[3]/div/button[2]")); }
        }

        public IWebElement PlansConfirmation
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tabs']/ul/li[3]/a")); }
        }

        public IWebElement RefreshButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='reload_plans']")); }
        }

        public IWebElement AcceptPlanButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='accept_plan']")); }
        }

        public IWebElement ConfirmPlanButton
        {
            get { return _firefox.FindElement(By.XPath("html/body/div[@aria-labelledby='ui-dialog-title-dialog-confirm']/div[3]/div/button[1]")); }
        }

        public IWebElement ApprovePlanButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='approve_plan']")); }
        }

        public IWebElement SpravochyeDannyeButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tabs']/ul/li[2]")); }
        }

        public IWebElement SalesData2016
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tab_info']/ul/li[1]")); }

        }

        public IWebElement SalesData2017
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tab_info']/ul/li[2]")); }

        }

        public IWebElement TotalSumSpravochyeDannye2016
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableprodprep_customer']/tbody/tr[1]/td[2]")); }
        }

        public IWebElement TotalSumSpravochyeDannye2017
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableprodprep_customer_cur_year']/tbody/tr[1]/td[2]")); }
           
        }

        public IWebElement AuditDataOwn2015
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tab_info']/ul/li[3]")); }
        }

        public IWebElement AuditDataCompetitor2015
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tab_info']/ul/li[5]")); }
        }

        public IWebElement TotalSumOwnSalesData2015
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableprodprep']/tbody/tr[1]/td[2]")); }
        }

        public IWebElement TotalSumCompetitorSalesData2015
        {
            get { return _firefox.FindElement(By.XPath(".//*[@id='tableprodconc']/tbody/tr[1]/td[2]")); }
        }
    }
}
