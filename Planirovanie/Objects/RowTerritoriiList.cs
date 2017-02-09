using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie
{
    public class RowTerritoriiList : List<RowTerritorii>
    {
        public List<string> GetTerritoriyList()
        {
            return (from row in this
                    select row.Name1RegionName).ToList();

        }

        public List<RowTerritorii> GetUniqueNoteBySeveralFields()
        {
            //return this.Select(r => r.WebName).Distinct().ToList();
           /* return (from r in this 
                    select r.WebName).Distinct().ToList();*/
            return (from r in this
                    group r by new {r.FIO, r.BuId, r.IdSotr} into grp
                    select grp.First()).ToList();
        }

        public static List<RowTerritorii> CompareTerritoriis(List<RowTerritorii> list1, List<RowTerritorii> list2) //Ищем в списке list1 то что не входит в list2
        {
            return (from data1 in list1 where !list2.Any(data2 => data2.FIO == data1.FIO && data2.Name1RegionName == data1.Name1RegionName) select data1).ToList();
        }
        public static List<RowTerritorii> CompareBuId(List<RowTerritorii> list1, List<RowTerritorii> list2) //Ищем в списке list1 то что не входит в list2
        {
            return (from data1 in list1 where !list2.Any(data2 => data2.FIO == data1.FIO && data2.BuId == data1.BuId && data1.IdSotr == data2.IdSotr) select data1).ToList();
        }
    }
}
