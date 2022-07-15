using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;


namespace Services.Services
{
    public interface IService
    {
        string ExtractPrecondition(string path, int i);

        //List<Configuration> ExtractConfiguration(string path, Worksheet xlWorkSheet);

        List<Step> ExtractStep(string path, int i);

        //TestCaseCaracteristics ExtractTestCaseCaracteristics(string path);

        TestDescriptionTable ExtractTestDescription(string path);

        DocumentVersionTable ExtractDocumentVersion(string path);

        ReferencesTable ExtractReferencesTable(string path);

        ValidationTable ExtractValidationTable(string path);

        List<TestCase> ExtractTestCases(string path);

        Scenario ExtractScenario(string path);

        FunctionalTestDocument ExtractGlobalFile(string path);
        


            string GetName();
    }
}
