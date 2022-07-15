using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services.Models.Json
{
    public class JsonStep
    {
        public string Seq { get; set; }

        // public string StepDescription { get; set; }
        public string Step { get; set; }

        public string ExpectedResult { get; set; }


        public bool IsEmpty()
        {
            if ((Seq == null || (Seq.Trim(' ')).Length == 0)
                && (Step == null || (Step.Trim(' ')).Length == 0)
                && (ExpectedResult == null || (ExpectedResult.Trim(' ')).Length == 0))
            {
                return true;
            }

            return false;
        }
    }
}
