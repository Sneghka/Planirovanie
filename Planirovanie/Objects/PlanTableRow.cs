﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planirovanie.Objects
{
    public class PlanTableRow
    {
        public string UserName { get; set; }
        public int UserId { get; set; }
        public string PreparationName { get; set; }
        public string TerritoriaName { get; set; }
        public int JanPsc { get; set; }
        public int FebPsc { get; set; }
        public int MarPsc { get; set; }
    }
}
