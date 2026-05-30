using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class PracticeArea
    {
        public EPAcode PAcode { get; set; }
        public ESampleType SampleType { get; set; } = ESampleType.not;
        public string Name { get; set; } // Full name of the practice area
        public string NameChinese { get; set; } // Full name of the practice area in Chinese
        public string Intent { get; set; } // Intent statement in English
        public string IntentChinese { get; set;  } // Intent statement in Chinese
        public string Value { get; set; } // Value statement in English
        public string ValueChinese { get; set; } // Value statement in Chinese

        public List<Practice> Practices { get; set; } = new List<Practice>(); // List of practiciese

        //public override string ToString()
        //{
        //    return $"{PAcode} - {Name} ({Practices.Count()})";
        //}
    }

    


}
