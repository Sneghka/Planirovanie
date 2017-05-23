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

        public static List<RowData> GetPreparationWithAutoPlanFromSpravochnik(List<RowData> spravochnik)
        {
           return (from l in spravochnik
                    where l.Group == "2 группа (автопланирование)"
                    select l).ToList();

        }
        public static List<RowData> GetPreparationWithAutoPlanFromPlanirovschik(RowDataList planirovschik)
        {
            return (from l in planirovschik
                    where l.Status == "false"
                    select l).ToList();

        }

        public static List<RowData> CompareRowDataObjects(List<RowData> list1, List<RowData> list2)
        {

            return (from l1 in list1
                    where !list2.Any(l1.IsEqual)
                    select l1).ToList();

            //list1.Where (r => !list2.Any (t => t.IsEqual( r ) ) )
        }

        public int GetTotalPcsById(int id, int[] months)
        {
            return (from r in this
                    where r.IdPrUniq == id && months.Contains(r.Month)
                    select r.Upakovki).Sum();
        }

        public int GetTotalPcsByIdAndUserAndMonths(int idPreparation, int[] months, int idSotr)
        {
            return (from r in this
                    where r.IdPrUniq == idPreparation && months.Contains(r.Month) && r.IdSotr == idSotr
                    select r.Upakovki).Sum();
        }

        public int GetTotalPcsByIdAndUser(int idPreparation, int userId)
        {
            return (from r in this
                    where r.IdPrUniq == idPreparation && r.IdSotr == userId
                    select r.Upakovki).Sum();
        }

        public decimal GetTotalSumByIdAndUser(int idPreparation, int userId)
        {
            return (from r in this
                    where r.IdPrUniq == idPreparation && r.IdSotr == userId
                    select r.Summa).Sum();
        }

        public decimal GetTotalSumRubById(int id, int[] months)
        {
            return (from r in this
                    where r.IdPrUniq == id && months.Contains(r.Month)
                    select r.Summa).Sum();
        }

        public decimal GetTotalSumRubByIdAndUserAndMonths(int id, int[] months, int idSotr)
        {
            return (from r in this
                    where r.IdPrUniq == id && months.Contains(r.Month) && r.IdSotr == idSotr
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

        public List<int> GetIdListByUserWithoutAutoplan(int user, int[] months)
        {
            return (from r in this
                    where r.IdSotr == user && r.Group != "2 группа (автопланирование)" && months.Contains(r.Month)
                    select r.IdPrUniq = r.Segment == 2 ? -r.IdPrUniq : r.IdPrUniq).Distinct().ToList();
        }

        public List<int> GetIdListByUserWithAutoplan(int user, int[] months)
        {
            return (from r in this
                    where r.IdSotr == user && months.Contains(r.Month)
                    select r.IdPrUniq = r.Segment == 2 ? -r.IdPrUniq : r.IdPrUniq).Distinct().ToList();
        }

        public List<int> GetIdListBySegmentWithoutAutoplan(int segment)
        {
            return (from r in this
                    where r.Segment == segment && r.Group != "2 группа (автопланирование)"
                    select r.IdPrUniq = r.Segment == 2 ? -r.IdPrUniq : r.IdPrUniq).Distinct().ToList();
        }

        public List<int> GetIdListBySegmentWithAutoplan(int segment, int[] months)
        {
            return (from r in this
                    where r.Segment == segment && months.Contains(r.Month)
                    select r.IdPrUniq = r.Segment == 2 ? -r.IdPrUniq : r.IdPrUniq).Distinct().ToList();
        }

        public List<int> GetIdList()
        {
            return (from r in this
                    select r.Segment == 2 ? r.IdPrUniq*-1 : r.IdPrUniq ).ToList();
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
                    select r.Upakovki = r.Id_BU == 105 ? r.Upakovki/2 : r.Upakovki).Sum();
        }

        public int GetUpakovkiByIdAndSalesType(int id, int salesType)
        {
            return (from r in this
                    where r.IdPrUniq == id && r.SalesType == salesType
                    select r.Upakovki = r.Id_BU == 105 ? r.Upakovki / 2 : r.Upakovki).Sum();
        }

        public int GetUpakovkiByIdAndTwoSalesType(int id, int salesType1, int salesType2)
        {
            return (from r in this
                    where r.IdPrUniq == id && (r.SalesType == salesType1 || r.SalesType == salesType2) && r.AreaName.Contains("Россия")
                    select r.Upakovki = r.Id_BU == 105 ? r.Upakovki / 2 : r.Upakovki).Sum();
        }

        public int GetUpakovkiConcurentById(int id)
        {
            return (from r in this
                    where r.IdPrUniq == id
                    select r.UpakovkiConcurent = r.Id_BU == 105 ? r.UpakovkiConcurent / 2 : r.UpakovkiConcurent).Sum();
        }

        public int GetUpakovkiConcurentByIdAndSalesType(int id, int salesType)
        {
            return (from r in this
                    where r.IdPrUniq == id && r.SalesType == salesType
                    select r.UpakovkiConcurent = r.Id_BU == 105 ? r.UpakovkiConcurent / 2 : r.UpakovkiConcurent).Sum();
        }

        public int GetUpakovkiConcurentByIdAndTwoSalesType(int id, int salesType1, int salesType2)
        {
            return (from r in this
                    where r.IdPrUniq == id && (r.SalesType == salesType1 || r.SalesType == salesType2)&&r.AreaName.Contains("Россия")
                    select r.UpakovkiConcurent = r.Id_BU == 105 ? r.UpakovkiConcurent / 2 : r.UpakovkiConcurent).Sum();
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

        public int GetUpakovkiAuditTmCommerciaOwn(int prepId, int oblId )
        {
            return (from r in this
                where r.IdPrUniq == prepId && r.DataType == 0 && r.SalesType == 1 && r.AreaId == oblId
                select r.Upakovki).Sum();
        }
        public int GetUpakovkiAuditTmCommerciaOwnLgota(int prepId, int oblId)
        {
            return (from r in this
                    where r.IdPrUniq == prepId && r.DataType == 0 && (r.SalesType == 2 || r.SalesType == 3) && r.AreaId == oblId
                    select r.Upakovki).Sum();
        }
        public int GetUpakovkiAuditTmCommerciaOwn67_115(int prepId, int oblId)
        {
            return (from r in this
                    where r.IdPrUniq == prepId && r.DataType == 0 && r.AreaId == oblId
                    select r.Upakovki).Sum();
        }

        public int GetUpakovkiAuditTmCommerciaConcurent(int prepId, int oblId)
        {
            return (from r in this
                    where r.IdPrUniq == prepId && r.DataType == 1 && r.SalesType == 1 && r.AreaId == oblId
                    select r.Upakovki).Sum();
        }
        public int GetUpakovkiAuditTmCommerciaConcurentLgota(int prepId, int oblId)
        {
            return (from r in this
                    where r.IdPrUniq == prepId && r.DataType == 1 && (r.SalesType == 2 || r.SalesType == 3) && r.AreaId == oblId
                    select r.Upakovki).Sum();
        }
        public int GetUpakovkiAuditTmCommerciaConcurent67_115(int prepId, int oblId)
        {
            return (from r in this
                    where r.IdPrUniq == prepId && r.DataType == 1 && r.AreaId == oblId
                    select r.Upakovki).Sum();
        }
    }
}
