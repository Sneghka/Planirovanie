using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
    public class DistributionSpravochnikRow
    {
        public int PreparationId { get; set; }
        public string PreparationName { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int? Segment { get; set; }
        public string RegionName1 { get; set; }
        public string DistrictName2 { get; set; }
        public int? SalesTypeId { get; set; }
        public int Upakovki { get; set; }
        public decimal Rubli { get; set; }
    }
}
