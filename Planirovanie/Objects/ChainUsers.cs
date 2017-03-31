using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{

   
    public class PmListId
    {
       public List<string> PmId { get; set; } = new List<string>();
    }

    public class NopListId
    {
        public List<int> NopId { get; set; } = new List<int>();
    }

    public class TmListId
    {
       public List<int> TmId { get; set; } = new List<int>();
       public List<string> RegionName { get; set; } = new List<string>();
    }

}
