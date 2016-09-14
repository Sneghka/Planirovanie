using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie
{
    public class RowTerritorii
    {
        public int IdSotr { get; set; }
        public string FIO { get; set; }
        public string Position { get; set; }
        public int DistrictId { get; set; }
        public string DistrictName { get; set; }
        public string BuId { get; set; }


        public bool AreRowTerritoriiEqual(RowTerritorii rowTerritorii)
        {
            return FIO == rowTerritorii.FIO && DistrictName == rowTerritorii.DistrictName;
        }


    }
}
