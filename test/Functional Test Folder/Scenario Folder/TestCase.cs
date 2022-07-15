namespace test
{
    using System.Collections.Generic;

    public class TestCase
    {
        public string Title { get; set; }

        public TestCaseCaracteristics TestCaseCaracteristics { get; set; }

        public string Context { get; set; }

        public string Aim { get; set; }

        public List<PreCondition> PreCondition { get; set; }

        public string PostCondition { get; set; }

        public List<Step> Step { get; set; }

        public string Repetition { get; set; }

        public List<Configuration> Configuration { get; set; }




    }
}
