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
    public static class CompareXlsFilesMethods
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
        public static void OutputMismatchFields(List<UniversalExcelObj> item1, List<UniversalExcelObj> item2, int i, int j)
        {
            Console.WriteLine("ClientID " + item1[i].ID + " - данные не совпадают");
            if (item1[i].Field1 != item2[j].Field1) Console.WriteLine("Не совпадает поле 1");
            if (item1[i].Field2 != item2[j].Field2) Console.WriteLine("Не совпадает поле 2");
            if (item1[i].Field3 != item2[j].Field3) Console.WriteLine("Не совпадает поле 3");
            if (item1[i].Field4 != item2[j].Field4) Console.WriteLine("Не совпадает поле 4");
            if (item1[i].Field5 != item2[j].Field5) Console.WriteLine("Не совпадает поле 5");
            if (item1[i].Field6 != item2[j].Field6) Console.WriteLine("Не совпадает поле 6");
            if (item1[i].Field7 != item2[j].Field7) Console.WriteLine("Не совпадает поле 7");
            if (item1[i].Field8 != item2[j].Field8) Console.WriteLine("Не совпадает поле 8");
            if (item1[i].Field9 != item2[j].Field9) Console.WriteLine("Не совпадает поле 9");
            if (item1[i].Field10 != item2[j].Field10) Console.WriteLine("Не совпадает поле 10");


        }

        public static void CompareAllFieldsById()
        {

            List<UniversalExcelObj> sortedFile_1 = FileXls_1.OrderByDescending(o => o.ID).ToList();
            List<UniversalExcelObj> sortedFile_2 = FileXls_2.OrderByDescending(o => o.ID).ToList();



            var minArrayLenght = sortedFile_1.Count < sortedFile_2.Count ? sortedFile_1.Count : sortedFile_2.Count;
            int j = 0;
            int lastIndexEquivalentInFile1 = 0;
            int lastIndexEquivalentInFile2 = 0;

            for (int i = 0; i < minArrayLenght;)
            {
                lastIndexEquivalentInFile1 = i;
                lastIndexEquivalentInFile2 = j;

                if (sortedFile_1[i].ID == sortedFile_2[j].ID)
                {
                   
                    if (!sortedFile_1[i].IsEqual(sortedFile_2[j]))
                    {
                        OutputMismatchFields(sortedFile_1, sortedFile_2, i, j);
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
                    Console.WriteLine("Элемент из второго списка - " + sortedFile_2[j].ID +
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

            if (lastIndexEquivalentInFile2 < sortedFile_2.Count-1)
            {
                for (int n = lastIndexEquivalentInFile2 + 1; n < sortedFile_2.Count; n++)
                {
                    Console.WriteLine("Элемент из второго списка - " + sortedFile_2[n].ID +
                                      " -  не содержится в первом списке");
                }
            }

            if (lastIndexEquivalentInFile1 < sortedFile_1.Count - 1)
            {
                for (int n = lastIndexEquivalentInFile1 + 1; n < sortedFile_1.Count; n++)
                {
                    Console.WriteLine("Элемент из первого списка - " + sortedFile_1[n].ID +
                                      " -  не содержится во вотором списке");
                }
            }



        }
    }
}
