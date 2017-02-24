using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.CheckStadaPlan
{
    public class User
    {
        public int BuId { get; set; }
        public int UserId { get; set; }
        public string UserName { get; set; }
        public List<int> TerritoryIdArray { get; set; } = new List<int>();
        public string Email { get; set; }
        public string Status { get; set; }
        
    }
}
