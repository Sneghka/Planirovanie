using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework.Constraints;
using Planirovanie.Objects;
using Excel = Microsoft.Office.Interop.Excel;

namespace Planirovanie.EcxelObj
{
    public static class CompareXlsFiles
    {
        private static UniversalExcelObjList FileXls_1 = new UniversalExcelObjList();
        private static UniversalExcelObjList FileXls_2 = new UniversalExcelObjList();

        public static void StoreExcelDataFromFileXls_1(string path)
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path, "Select * from [Лист1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new UniversalExcelObj()
                {
                    ID = Convert.ToInt32(row["DicClientID"]),
                    /* ID =row["DicClientID"].ToString(),*/
                    Field1 = row["Name"].ToString(),
                    Field2 = row["CleanAddress"].ToString(),
                    Field3 = row["Building"].ToString(),
                    Field4 = row["KLADRid"].ToString(),
                    Field5 = row["Index"].ToString(),
                    Field6 = row["RegionID"].ToString(),
                    Field7 = row["ClientType_Code"].ToString(),
                    Field8 = row["PharmacyNet_Code"].ToString(),
                    Field9 = row["INN"].ToString(),
                    Field10 = row["LegalName"].ToString()
                };
                FileXls_1.Add(rowData);
            }
        }
        public static void StoreExcelDataFromFileXls_2(string path)
        {
            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, path, "Select * from [Лист1$]");
            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;
                var rowData = new UniversalExcelObj()
                {
                    ID = Convert.ToInt32(row["DicClientID"]),
                    /* ID = row["DicClientID"].ToString(),*/
                    Field1 = row["Name"].ToString(),
                    Field2 = row["CleanAddress"].ToString(),
                    Field3 = row["Building"].ToString(),
                    Field4 = row["KLADRid"].ToString(),
                    Field5 = row["Index"].ToString(),
                    Field6 = row["RegionID"].ToString(),
                    Field7 = row["ClientType_Code"].ToString(),
                    Field8 = row["PharmacyNet_Code"].ToString(),
                    Field9 = row["INN"].ToString(),
                    Field10 = row["LegalName"].ToString()
                };
                FileXls_2.Add(rowData);
            }
        }

        public static void CompareAllFieldsById()
        {

            List<UniversalExcelObj> sortedFile_1 = FileXls_1.OrderByDescending(o => o.ID).ToList();
            List<UniversalExcelObj> sortedFile_2 = FileXls_2.OrderByDescending(o => o.ID).ToList();

            /*   for (int i = 0; i < 10; i++)
               {
                   Console.WriteLine(sortedFile_1[i].ID + " / " + sortedFile_2[i].ID);
               }*/

            for (int i = 0; i < sortedFile_1.Count;)
            {
                for (int j = 0; j < sortedFile_2.Count;)
                {
                    if (sortedFile_1[i].ID == sortedFile_2[j].ID)
                    {
                        if (!sortedFile_1[i].IsEqual(sortedFile_2[j]))
                        {
                            Console.WriteLine("ClientID " + sortedFile_1[i].ID + " - данные не совпадают");
                            i++;
                            j++;
                            continue;
                        }
                        i++;
                        j++;
                        continue;

                    }

                    if (sortedFile_1[i].ID < sortedFile_2[j].ID)
                    {
                        Console.WriteLine("Элемент из второго списка - " + sortedFile_2[i].ID +
                                          " -  не содержится в первом списке");
                        j++;
                    }
                    if (sortedFile_1[i].ID > sortedFile_2[j].ID)
                    {
                        Console.WriteLine("Элемент из первого списка - " + sortedFile_1[i].ID +
                                          " -  не содержится во вотором списке");
                        i++;
                    }
                }
            }
        }
    }
}
