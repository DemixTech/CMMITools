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
        public int Level { get; set; } // Practice level, 1..5
        public int Number { get; set; } // Practice level number for exmaple 1.1 or 1.2 (the 1 or 2)
        public string Statement { get; set; } // Practice statement
        public string StatementChinese { get; set; } // Practice statement in Chinese

        public List<Question> Questions { get; set; } = new List<Question>(); // List of quetions 
        public List<ExampleActivity> ExampleActivities { get; set; } = new List<ExampleActivity>();
        public List<ExampleWorkProduct> ExampleWorkProducts { get; set; } = new List<ExampleWorkProduct>();
        public override string ToString()
        {
            return $"{Acronym} {Level}.{Number}";
        }
    }
}
