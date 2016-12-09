using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie
{
    public class RowDataList : List<RowData>
    {
        public List<string> GetUniqueWebNames(int[] months)
        {
            //return this.Select(r => r.WebName).Distinct().ToList();
            return (from r in this
                    where r.Group != "не планируем в Планировщике" && months.Contains(r.Month)
                    select r.WebName).Distinct().ToList();
        }
        public List<string> GetUniqueNames()
        {
            //return this.Select(r => r.WebName).Distinct().ToList();
            return (from r in this
                    select r.Name).Distinct().ToList();
        }

       
        public static List<RowData> ConvertSpravochikList(int[] months, RowDataList spravochnik)
        {
            var clearedList = spravochnik.Select(r => new RowData { IdPrUniq = r.Segment == 2 ? -r.IdPrUniq : r.IdPrUniq, Name = r.Segment == 2 ? r.WebName : r.Name, Segment = r.Segment, Id_BU = r.Id_BU, Group = r.Group, Month = r.Month}).Where(r=> r.Group != "не планируем в Планировщике" && months.Contains(r.Month))
                .Distinct().ToList();
            var finalList = clearedList.GroupBy(r => new { r.IdPrUniq, r.Name, r.Segment, r.Id_BU, r.Group })
                 .Select(g => g.First())
                 .ToList();

            return finalList;
        }

        public static List<RowData> CompareRowDataObjects(List<RowData> list1, List<RowData> list2)
        {

            return (from l1 in list1
                    where !list2.Any(l1.IsEqual)
                    select l1).ToList();

            //list1.Where (r => !list2.Any (t => t.IsEqual( r ) ) )
        }

        public int GetTotalSumPcsById(int id, int[] months)
        {
            return (from r in this
                    where r.IdPrUniq == id && months.Contains(r.Month)
                    select r.Upakovki).Sum();
        }
        public decimal GetTotalSumRubById(int id, int[] months)
        {
            return (from r in this
                    where r.IdPrUniq == id && months.Contains(r.Month)
                    select r.Summa).Sum();
        }


        public List<int> GetIdListByUser(int user)
        {
            return (from r in this
                    where r.IdSotr == user
                    select r.IdPrUniq).Distinct().ToList();
        }

        public List<int> GetIdListByUserWithoutAutoplan(int user)
        {

           
            return   (from r in this
                    where r.IdSotr == user && r.Group != "2 группа (автопланирование)"
                      select r.IdPrUniq = r.Segment == 2 ? -r.IdPrUniq : r.IdPrUniq).Distinct().ToList();
        }

        public List<int> GetIdList()
        {
            return (from r in this
                    select r.IdPrUniq).ToList();
        }

        public static List<string> CompareStrings(List<string> list1, List<string> list2)
        {
            //return list1.Where(s => !list2.Contains(s)).ToList();
            return (from s in list1
                    where !list2.Contains(s)
                    select s).ToList();
        }
        public static List<int> CompareNumbers(List<int> list1, List<int> list2)
        {

            return (from s in list1
                    where !list2.Contains(s)
                    select s).ToList();
        }


        public List<RowData> GetListObjectsById(int id)
        {
            return (from r in this
                    where r.IdPrUniq == id
                    select r).ToList();
        }
        public int GetSumPcsByIdAndMonth(int id, int month)
        {
            return (from r in this
                    where r.IdPrUniq == id && r.Month == month
                    select r.Upakovki).Sum();
        }
        public int GetPcsByIdAndSegmentAndMonth(int id, int month, int segment)
        {
            return (from r in this
                    where r.IdPrUniq == id && r.Month == month && r.Segment == segment
                    select r.Upakovki).Sum();
        }



        public int GetSumByChoosenMonth(int month)
        {
            return (from r in this
                    where r.Month == month
                    select r.Upakovki).Sum();
        }

        public int GetUpakovkiById(int id)
        {
            return (from r in this
                    where r.IdPrUniq == id
                    select r.Upakovki).Sum();
        }
        public int GetUpakovkiConcurentById(int id)
        {
            return (from r in this
                    where r.IdPrUniq == id
                    select r.UpakovkiConcurent).Sum();
        }


        public List<string> GetNamesList()
        {
            return (from row in this
                    select row.Name).ToList();

        }

        public List<string> GetWebNamesList()
        {
            return (from row in this
                    select row.WebName).ToList();

        }
    }
}
