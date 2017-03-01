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
                    group r by new { r.FIO, r.BuId, r.IdSotr } into grp
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


        public List<int> GetRaionTerritorr(int userId)
        {
            return (from row in this
                    where row.IdSotr == userId && row.Name3RaionId != string.Empty
                    select Convert.ToInt32(row.Name3RaionId)).Distinct().ToList();
        }

        public List<int> GetRaionTerritorrByOblastIdAndUserId(int oblastId, int userId)
        {
            return (from row in this
                    where row.Name3RaionId != string.Empty && Convert.ToInt32(row.Name2OblastId) == oblastId && Convert.ToInt32(row.IdSotr) == userId && Convert.ToInt32(row.Name1RegionId) != 8331 // БЕЗ КРЫМА
                    select Convert.ToInt32(row.Name3RaionId)).Distinct().ToList();
        }
        public List<int> GetOblastTerritorr(int userId)
        {
            return (from row in this
                    where row.IdSotr == userId && row.Name2OblastId != string.Empty
                    select Convert.ToInt32(row.Name2OblastId)).Distinct().ToList();
        }


        public List<int> GetRegionTerritorr(int userId)
        {
            return (from row in this
                    where row.IdSotr == userId && row.Name1RegionId != string.Empty && row.BuId != 67 && Convert.ToInt32(row.Name1RegionId) != 8331 // БЕЗ КРЫМА
                    select Convert.ToInt32(row.Name1RegionId)).Distinct().ToList();
        }

        public List<int> GetRegionTerritorr(int userId, int buNumber)
        {
            return (from row in this
                    where row.IdSotr == userId && row.Name1RegionId != string.Empty && row.BuId == buNumber && Convert.ToInt32(row.Name1RegionId) != 8331 // БЕЗ КРЫМА
                    select Convert.ToInt32(row.Name1RegionId)).Distinct().ToList();
        }

        public List<int> GetOblastterritorrByRegionId(int regionId)
        {
            return (from row in this
                    where row.Name2OblastId != string.Empty && Convert.ToInt32(row.Name1RegionId) == regionId && Convert.ToInt32(row.Name1RegionId) != 8331 // БЕЗ КРЫМА
                    select Convert.ToInt32(row.Name2OblastId)).Distinct().ToList();
        }
        public List<int> GetOblastterritorrByRegionIdUserBuId(int regionId, int userBuId)
        {
            return (from row in this
                    where row.Name2OblastId != string.Empty && Convert.ToInt32(row.Name1RegionId) == regionId && row.BuId == userBuId && Convert.ToInt32(row.Name1RegionId) != 8331 // БЕЗ КРЫМА
                    select Convert.ToInt32(row.Name2OblastId)).Distinct().ToList();
        }


        public List<int> GetUserIdListByBuId(int buId)
        {
            return (from row in this
                    where row.BuId == buId
                    select row.IdSotr).Distinct().ToList();
        }

        public bool IsUserExistInSpravochink(int userId)
        {
            foreach (var row in this)
            {
                if (row.IdSotr == userId) return true;
            }
            return false;
        }

        public bool IsBuUserSpravochikMatchPlanirovschik(int buUserIdPlanirovschik, int userId)
        {
            var getBuSpravochnik = (from r in this
                                    where r.IdSotr == userId
                                    select r.BuId).Distinct().ToList();

           if(getBuSpravochnik.Count==1) return true;
            return false;
        }

    }
}
