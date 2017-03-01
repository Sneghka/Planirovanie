using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
    public static class PlanTableRowList
    {

        public static int GetTotalPcsByMonth(List<PlanTableRow> plan, int month)
        {
            return (from row in plan
                    where row.Month == month
                    select row.Pcs).Sum();
        }

        public static bool IsTotalMatch(List<PlanTableRow> plan1340, List<PlanTableRow> planUser)
        {
            if (GetTotalPcsByMonth(plan1340, 1) != GetTotalPcsByMonth(planUser, 1) || GetTotalPcsByMonth(plan1340, 2) != GetTotalPcsByMonth(planUser, 2) || GetTotalPcsByMonth(plan1340, 3) != GetTotalPcsByMonth(planUser, 3)) return false;
            return true;
        }

        public static List<string> ComparePreparationName(List<PlanTableRow> plan1340, List<PlanTableRow> planUser)
        {
            var diffList = new List<string>();
            var prep1340 = (from r in plan1340
                select r.PreparationName).Distinct().ToList();

            var prepUser = (from r in planUser
                            select r.PreparationName).Distinct().ToList();

            var notExistInUserPlan = (from name in prep1340
                where !prepUser.Contains(name)
                select name + " отсутствует в плане пользователя").ToList();

            var notExistIn1340Plan = (from name in prepUser
                                      where !prep1340.Contains(name)
                                      select name + " отсутствует в плане 1340").ToList();
            diffList.AddRange(notExistIn1340Plan);
            diffList.AddRange(notExistInUserPlan);
            return diffList;
        }


        public static void ComparePlans(List<PlanTableRow> plan1340, List<PlanTableRow> planUser, List<PlanTableRow> planBu33)
        {
            foreach (var row in plan1340)
            {

                var diff = (from rowUser in planUser
                            where
                                row.PreparationName == rowUser.PreparationName && row.Month == rowUser.Month &&
                                row.TerritoriaName == rowUser.TerritoriaName && row.Pcs != rowUser.Pcs
                            select rowUser).ToList();

                foreach (var d in diff)
                {

                    var pcsBU33 = (from r in planBu33
                                   where d.PreparationName == r.PreparationName && d.Month == r.Month &&
                                          d.TerritoriaName == r.TerritoriaName
                                   select r.Pcs).Sum();

                    
                    if (d.Pcs + pcsBU33 != row.Pcs)
                    {
                        Console.WriteLine("Упаковки не совпадают");
                        Console.WriteLine(d.PreparationName + " " + d.TerritoriaName + " " + " месяц " + d.Month + row.Pcs + "(план 1340) " + d.Pcs + "(план пользователь) " + pcsBU33 + "(план БЮ33)");
                    }

                }
            }

            Console.WriteLine("  План проверен");
        }
    }
}
