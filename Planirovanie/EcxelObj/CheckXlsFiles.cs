using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium.Firefox;
using Planirovanie.Objects;


namespace Planirovanie.EcxelObj
{
    [TestFixture]
   public class CheckXlsFiles
    {

        const string path_file1 = @"D:\Sneghka\Selenium\Projects\Files\alg_new.xlsx";
        const string path_file2 = @"D:\Sneghka\Selenium\Projects\Files\alg_old.xlsx";
        /*const string path_file1 = @"D:\Sneghka\Selenium\Projects\Files\alg_new1.xlsx";
        const string path_file2 = @"D:\Sneghka\Selenium\Projects\Files\alg_old1.xlsx";*/
        [Test]
        public void CompareFieldsInExlFiles()
        {
            CompareXlsFiles.StoreExcelDataFromFileXls_1(path_file1);
            Console.WriteLine("Store data from fist file");
            CompareXlsFiles.StoreExcelDataFromFileXls_2(path_file2);
            Console.WriteLine("Store data from second file");
            CompareXlsFiles.CompareAllFieldsById();
          

        }
    }
}
