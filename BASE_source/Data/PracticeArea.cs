using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class PracticeArea
    {
        public EPAcode PAcode;
        public ESampleType SampleType = ESampleType.not;
        
        public string Name { get; set; } // Full name of the practice area
        public string NameChinese { get; set; } // Full name of the practice area in Chinese
        public string Intent { get; set; } // Intent statement in English
        public string IntentChinese { get; set;  } // Intent statement in Chinese
        public string Value { get; set; } // Value statement in English
        public string ValueChinese { get; set; } // Value statement in Chinese

        public List<Practice> Practices { get; set; } = new List<Practice>(); // List of practiciese

        public override string ToString()
        {
            return $"{PAcode} - {Name} ({Practices.Count()})";
        }
    }

    public enum EPAcode
    {
        PI = 1,
        TS = 2,
        PQA = 3,
        PR = 4,
        RDM = 5,
        VV = 6,
        MPM = 7,
        PAD = 8,
        PCM = 9,
        RSK = 10,
        OT = 11,
        EST = 12,
        MC = 13,
        PLAN = 14,
        CAR = 15,
        CM = 16,
        DAR = 17,
        SAM = 18,
        II = 19,
        GOV = 20,

    }

    public enum ESampleType
    {
        not = 0, // The practice is not performed. Don't list it then. There should not be .not
        sampled = 1, // the practice is sampled by the generator
        added = 2, // the practice is added by the LA
        perform = 3, // the practice is performed by the Participant
        model =4, // the practice is related to a model structure. Set this type 
    }

}
