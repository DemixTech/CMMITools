using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public enum ESampleType
    {
        not = 0, // The practice is not performed. Don't list it then. There should not be .not
        sampled = 1, // the practice is sampled by the generator
        added = 2, // the practice is added by the LA
        perform = 3, // the practice is performed by the Participant
        model = 4, // the practice is related to a model structure. Set this type 
    }
}
