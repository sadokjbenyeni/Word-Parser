using System.Collections.Generic;

namespace Services.Services
{
    public class Scenario
    {

        public string Context { get; set; }
        
        public string Type { get; set; }

        public string Target { get; set; }

        public string ResearchOrientation { get; set; }

        public string Environment { get; set; }

        public List<TestCase> TestCases { get; set; }



    }
}
