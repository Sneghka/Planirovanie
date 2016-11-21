using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
    public class UniversalExcelObjList: List<UniversalExcelObj>
    {
        public static List<UniversalExcelObj> CompareXlsDataSmallObjects(UniversalExcelObjList file1, UniversalExcelObjList file2)
        {

            return (from l1 in file1
                    where !file2.Any(l1.IsEqual)
                    select l1).ToList();

            //list1.Where (r => !list2.Any (t => t.IsEqual( r ) ) )
        }
        public static List<UniversalExcelObj> GetItemById(UniversalExcelObjList file, int id)
        {
            return(from f in file
                where f.ID == id
                select f).ToList();
        }
    }



}


