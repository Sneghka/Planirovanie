using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
    public class UniversalExcelObj
    {

        public int ID { get; set; }
        public string Field1 { get; set; }
        public string Field2 { get; set; }
        public string Field3 { get; set; }
        public string Field4 { get; set; }
        public string Field5 { get; set; }
        public string Field6 { get; set; }
        public string Field7 { get; set; }
        public string Field8 { get; set; }
        public string Field9 { get; set; }
        public string Field10 { get; set; }
        public string Field11 { get; set; }



        public bool IsEqual(UniversalExcelObj anotherOne)
        {
            if (ID==anotherOne.ID &&
                Field1 == anotherOne.Field1 &&
                Field2 == anotherOne.Field2 &&
                Field3 == anotherOne.Field3 &&
                Field4 == anotherOne.Field4 &&
                Field5 == anotherOne.Field5 &&
                Field6 == anotherOne.Field6 &&
                Field7 == anotherOne.Field7 &&
                Field8 == anotherOne.Field8 &&
                Field9 == anotherOne.Field9 &&
                Field10 == anotherOne.Field10
                )
            {
                return true;
            }
            else
            {
                return false;
            }

        }
    }
}
