using Services.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Services.Models.Json
{
    public class JsonModelEmpty
    {

        public string Title { get; set; }

        public bool Active { get; set; }

        public string Description { get; set; }

        public string Status { get; set; }

        public string TestType { get; set; }

        public string Priority { get; set; }

        public int OwnerUserId { get; set; }

        public int FolderId { get; set; }

        public List<CustomFields> CustomFields { get; set; }
        

        public static JsonModelEmpty BuildInstance(FunctionalTestDocument globalObject, int iCase, int authorId, int dossierId,
            string Title, bool Active, string Description, string Status, string TestType, string Priority, List<JsonStep> JsonStepList, List<CustomFields> CustomList)
        {


            if (JsonStepList != null && JsonStepList.Count > 0 && JsonStepList.Any())
            {
                return new JsonModel()
                {
                    Title = globalObject.Scenario?.TestCases[iCase]?.Title,

                    Active = true,

                    Description = "<b>Context:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.Context
                  + " <br />" + " <br />" + "<b>Aim:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.Aim
                  + " <br />" + " <br />" + "<b>PreCondition:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.PreCondition
                  + " <br />" + " <br />" + "<b>PostCondition:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.PostCondition,

                    //TestType = globalObject.Scenario?.TestCases[iCase]?.TestCaseCaracteristics?.TypeQC.TrimStart(),

                    Priority = globalObject.Scenario?.TestCases[iCase]?.TestCaseCaracteristics?.Priority.TrimStart(),

                    OwnerUserId = authorId,

                    FolderId = dossierId,

                    Status = "New",

                    CustomFields = CustomList,

                    TestSteps = JsonStepList
                    


                };
            }
            else
            {
                return new JsonModelEmpty()
                {
                    Title = globalObject.Scenario?.TestCases[iCase]?.Title,

                    Active = true,

                    Description = "<b>Context:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.Context
                                      + " <br />" + " <br />" + "<b>Aim:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.Aim
                                      + " <br />" + " <br />" + "<b>PreCondition:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.PreCondition
                                      + " <br />" + " <br />" + "<b>PostCondition:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.PostCondition,

                    //TestType = globalObject.Scenario?.TestCases[iCase]?.TestCaseCaracteristics?.TypeQC.TrimStart(),

                    Priority = globalObject.Scenario?.TestCases[iCase]?.TestCaseCaracteristics?.Priority.TrimStart(),

                    OwnerUserId = authorId,

                    FolderId = dossierId,

                    Status = "New",

                    CustomFields = CustomList



                };
            }
        }
    }
}
