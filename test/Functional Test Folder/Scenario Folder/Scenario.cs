using System.Collections.Generic;

namespace test
{
    public class Scenario
    {
        public Scenario()
        {
            TestCases = new List<TestCase>();
        }

        public string Context { get; set; }
        
        public string Type { get; set; }

        public string Target { get; set; }

        public string ResearchOrientation { get; set; }

        public string Environment { get; set; }

        List<TestCase> TestCases { get; set; }



    }
}
