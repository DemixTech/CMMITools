using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class MapRecord
    {
        public string LevelStr { get; set; }
        public string PAstr { get; set; }
        public string PALevelStr { get; set; } // { return PAstr + " " + LevelStr; } }
        public int Row { get; set; }
        public int Col { get; set; }
        public string RowColStr { get; set; }
        public bool OoS { get; set; } // true if out of scope, false if in scope

    }
}
