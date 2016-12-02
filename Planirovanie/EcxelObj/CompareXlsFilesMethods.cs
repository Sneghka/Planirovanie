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

            var lengthFile1 = sortedFile_1.Count;
            var lengthFile2 = sortedFile_2.Count;

            int x = 0;
            int y = 0;

            while (x < lengthFile1 || y < lengthFile2)
            {

                if (x < lengthFile1 && y < lengthFile2 && sortedFile_1[x].ID < sortedFile_2[y].ID)
                {
                    Console.WriteLine("Элемент из второго списка - " + sortedFile_2[y].ID +
                                      " -  не содержится в первом списке");
                    y++;
                    continue;
                }
                if (x < lengthFile1 && y < lengthFile2 && sortedFile_1[x].ID > sortedFile_2[y].ID)
                {
                    Console.WriteLine("Элемент из первого списка - " + sortedFile_1[x].ID +
                                      " -  не содержится во вотором списке");
                    x++;
                    continue;
                }
                if (x < lengthFile1 && y < lengthFile2 && sortedFile_1[x].ID == sortedFile_2[y].ID)
                {
                    if (!sortedFile_1[x].IsEqual(sortedFile_2[y]))
                    {
                        OutputMismatchFields(sortedFile_1, sortedFile_2, x, y);
                    }
                    x++;
                    y++;
                    continue;
                }

                if (y >= lengthFile2)
                {
                    Console.WriteLine("Элемент из первого списка - " + sortedFile_1[x].ID +
                                      " -  не содержится во втором списке");
                    x++;
                    continue;
                }
                if (x >= lengthFile1)
                {
                    Console.WriteLine("Элемент из второго списка - " + sortedFile_2[y].ID +
                                     " -  не содержится в первом списке");
                    y++;
                }
            }
        }
    }
}
