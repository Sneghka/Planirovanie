using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
   public class DistributionSpravochnikRowList : List<DistributionSpravochnikRow>
    {
        public int GetUpakovkiByIdWithoutCrimea(int id)
        {
            return (from r in this
                    where r.PreparationId == id && r.RegionName1!= "Крым"
                    select r.Upakovki).Sum();
        }

        public int GetUpakovkiByIdBySegmentBySalesTypeWithoutCrimea(int id)
        {
            return (from r in this
                    where r.PreparationId == id && r.RegionName1 != "Крым" && (r.Segment == 1 || r.SalesTypeId == 1)
                    select r.Upakovki).Sum();
        }
       
        public int GetUpakovkiByIdBySegmentBySalesTypeWithoutCrimeaLgotaBU33(int id)
        {
            return (from r in this
                    where r.PreparationId == id && r.RegionName1 != "Крым" && (r.Segment == 2 || r.Segment == 3 || r.SalesTypeId == 2)
                    select r.Upakovki).Sum();
        }

      
        public int GetUpakovkiByIdBySegmentBySalesTypeByRegion(int id, string region)
        {
            return (from r in this
                    where r.PreparationId == id && r.RegionName1 != "Крым" && (r.Segment == 1 || r.SalesTypeId == 1) && r.RegionName1 == region
                    select r.Upakovki).Sum();
        }
    }
}
