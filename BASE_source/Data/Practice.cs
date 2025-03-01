using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class Practice
    {
        public string Acronym { get; set; }  // Practice acronum, for exmampel GOV
        //public string CodeAndNumber { get; set; }
        public int Level { get; set; } // Practice level, 1..5
        public int Number { get; set; } // Practice level number for exmaple 1.1 or 1.2 (the 1 or 2)
        public string Statement { get; set; } // Practice statement
        public string StatementChinese { get; set; } // Practice statement in Chinese

        public List<Question> Questions { get; set; } = new List<Question>(); // List of quetions 
        public List<ExampleActivity> ExampleActivities { get; set; } = new List<ExampleActivity>();
        public List<ExampleWorkProduct> ExampleWorkProducts { get; set; } = new List<ExampleWorkProduct>();
        //public List<WorkUnit> WorkUnits { get; set; } = new List<WorkUnit>();

        //public string weaknessStr { get; set; } = string.Empty;
        //public string strengthStr {  get; set; } = string.Empty;
        //public string recommendationStr { get; set; } = string.Empty;
        //public string charStr { get; set; } = string.Empty;
        //public string sessionStr { get; set; } = string.Empty;
        //public string participantsStr {  get; set; } = string.Empty;    


        //public override string ToString()
        //{
        //    return $"{Acronym} {Level}.{Number}";
        //}
    }
}
