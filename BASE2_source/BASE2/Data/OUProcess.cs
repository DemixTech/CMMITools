using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE2.Data
{
    public class OUProcess
    {
        public String Name { get; set; } // The name of the process
        public List<WorkUnit> WorkUnits { get; set; } = new List<WorkUnit>(); // The work units associated with this process
    }
}
