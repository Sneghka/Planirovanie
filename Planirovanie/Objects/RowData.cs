using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie
{
    public class RowData
    {
        public int IdPrUniq { get; set; }
        public int IdSotr { get; set; }
        public string Name { get; set; }
        public string FIO { get; set; }
        public int Id_BU { get; set; }
        public string BusinessUnit { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int Segment { get; set; }
        public int Upakovki { get; set; }
        public decimal Summa { get; set; }
        public string Group { get; set; }
        public int UpakovkiConcurent { get; set; }

        public string WebName
        {
            get { return Segment == 2 ? Name + " (льгота)" : Name; }
        }

        public bool IsEqual(RowData anotheRowData)
        {
           if (Name == anotheRowData.Name && IdPrUniq == anotheRowData.IdPrUniq) return true;
            return false;

        }
       
    }
}
