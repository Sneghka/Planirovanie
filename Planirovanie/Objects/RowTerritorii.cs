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
        public int BuId { get; set; }
        public string Name1RegionId { get; set; }
        public string Name1RegionName { get; set; }
        public string Name2OblastId { get; set; }
        public string Name2OblastName { get; set; }
        public string Name3RaionId { get; set; }
        public string Name3RaionName { get; set; }
        public string Email { get; set; }
        public string Login { get; set; }



        public bool AreRowTerritoriiEqual(RowTerritorii rowTerritorii)
        {
            return FIO == rowTerritorii.FIO && Name1RegionName == rowTerritorii.Name1RegionName;
        }


    }
}
