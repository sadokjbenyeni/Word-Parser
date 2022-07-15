
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Services.Models;
using System.Text;
using System.IO;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Web.Script.Serialization;
using Services.Models.Tables_Folder;
using System.ComponentModel;
using Forms = System.Windows.Forms;
using System.Data;
using Services.Models.Json;

namespace Services.Services
{
    public class Service : IService
    {
        public Application Application { get; set; }

        SortedDictionary<int, int> MapTestCaseSteps = new SortedDictionary<int, int>();


        public Service()
        {
            this.Application = new Application();
        }

        //*********************************************************************//Extract Configuration Table Method//*******************************************************************//
        public List<Configuration> ExtractConfiguration(string path, Excel.Worksheet xlWorkSheet)
        {
             short tableIndex = 15;
            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);

            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Instanciate a list of Configuration object  
            List<Configuration> configurationList = new List<Configuration>();

            //Define the table that we will be working on
            Table Table15 = allDocument.Tables[tableIndex];
           
            //Create a matrix that holds the same informations as in the Configuration Table
            string[,] ParameterMatrix = new string[Table15.Rows.Count, Table15.Columns.Count];


            //Loop Over Configration Table Columns
            for (int col = 1; col <= Table15.Columns.Count; col++)
            {
                //Instanciate Configuration object to fill it every time we go over a column
                Configuration configuration = new Configuration();

                //Loop Over Configuration Table Rows 
                for (int row = 1; row <= Table15.Rows.Count; row++)
                {
                    //Declare a Cell for as specific column and row
                    var parameterCell = Table15.Cell(row, col);

                    //Fill Parameter Matrix with the current cell value
                    ParameterMatrix[row - 1, col - 1] = ReplaceNonPrintableCharacters(parameterCell.Range.Text);

                    //Fill Configuration object with the previous value
                    configuration.Parameter = ParameterMatrix[row - 1, col - 1];

                    //Filling the Worksheet Cells with the data from ParameterMatrix Cells
                    xlWorkSheet.Cells[row, col] = ParameterMatrix[row - 1, col - 1];

                }


                //Adding all configuration table rows together
                configurationList.Add(configuration);

            }


            return configurationList;
        }


        //**********************************************************************//Extract PreCondtion Table Method//*******************************************************************//
        public string ExtractPrecondition(string path, int i)
        {

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);

            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Instanciate a list of Configuration object  
            List<PreCondition> preConditionList = new List<PreCondition>();

            //Define the table that we will be working on

            Table Table13 = allDocument.Tables[i];
            //Declaring a matrix with the same measurment of this table
            string[,] PreConditionMatrix = new string[Table13.Rows.Count - 1, Table13.Columns.Count];
            MyTable preTable = new MyTable();
            preTable.Columns.Add("Test Name", typeof(string));
            preTable.Columns.Add("Parameter", typeof(string));
            preTable.Columns.Add("Comments", typeof(string));
            Forms.DataGridView dataGridView = new Forms.DataGridView();

            //Going through the doc table and fill the matrix in the same order
            for (int row = 2; row <= Table13.Rows.Count; row++)
            {
                //Instanciate Precondition object to fill it every time we go over a column
                PreCondition preCondition = new PreCondition();

                //Loop Over Precondition Table Rows 
                for (int col = 1; col <= Table13.Columns.Count; col++)
                {
                    //Declare a Cell for as specific column and row
                    var preConditionCell = Table13.Cell(row, col);

                    //Fill Parameter Matrix with the current cell value
                    PreConditionMatrix[row - 2, col - 1] = ReplaceNonPrintableCharacters(preConditionCell.Range.Text);

                    // If statement is for choosing the first Column
                    if (col % 2 != 0 && col % 3 != 0)
                    {
                        preCondition.TestName = PreConditionMatrix[row - 2, col - 1];

                    }
                    //Else If statement is for selecting ony the second Column
                    else if (col % 2 == 0)
                    {
                        preCondition.Parameter = PreConditionMatrix[row - 2, col - 1];

                    }
                    //Else select the last Column
                    else
                    {
                        preCondition.Comments = PreConditionMatrix[row - 2, col - 1];

                    }


                }
                preTable.Rows.Add(preCondition.TestName, preCondition.Parameter, preCondition.Comments);
                //Add the preCondition object to the List of object to get a full table
                preConditionList.Add(preCondition);
                dataGridView.DataSource = preTable;


            }

            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<table border='1px' cellpadding='30' cellspacing='0' style='font-family:Garamond; font-size:smaller'>");

            strHTMLBuilder.Append("<tr>");
            strHTMLBuilder.Append("<tr>");
            foreach (DataColumn myColumn in preTable.Columns)
            {
                strHTMLBuilder.Append("<td bgcolor='#ffffcc'>");
                strHTMLBuilder.Append("<b>");
                strHTMLBuilder.Append(myColumn.ColumnName);
                strHTMLBuilder.Append("</b>");
                strHTMLBuilder.Append("</td>");



            }
            strHTMLBuilder.Append("</tr>");

            foreach (DataRow myRow in preTable.Rows)
            {

                strHTMLBuilder.Append("<tr>");
                foreach (DataColumn myColumn in preTable.Columns)
                {
                    strHTMLBuilder.Append("<td>");
                    strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");

                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</table>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;
        }

        
        //*****************************************************************************//Extract Step Method//*************************************************************************//
        public List<Step> ExtractStep(string path, int i)
        {

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);

            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);
            
            //Instanciate a list of Steps object  
            List<Step> stepList = new List<Step>();

            //Define the table that we will be working on
            Table Table14 = allDocument.Tables[i];

            //Create a matrix identic to the Step Table
            string[,] StepMatrix = new string[Table14.Rows.Count - 1, Table14.Columns.Count];

            //Loop Over Steps Table Rows 
            for (int row = 2; row <= Table14.Rows.Count; row++)
            {
                //Instanciate Steps object to fill it every time we go over a column
                Step step = new Step();

                //Loop Over Steps Table Rows 
                for (int col = 1; col <= Table14.Columns.Count; col++)
                {
                    //Declare a Cell for as specific column and row
                    var stepCell = Table14.Cell(row, col);

                    //Fill Parameter Matrix with the current cell value
                    StepMatrix[row - 2, col - 1] = ReplaceNonPrintableCharacters(stepCell.Range.Text);

                    // If statement is for choosing the first Column
                    if (col % 2 != 0 && col % 3 != 0)
                    {
                        step.StepNum = StepMatrix[row - 2, col - 1];

                    }

                    //Else If statement is for selecting ony the second Column
                    else if (col % 2 == 0 && col % 4 != 0)
                    {
                        step.Description = StepMatrix[row - 2, col - 1];

                    }

                    //Else select the third Column
                    else if (col % 3 == 0)
                    {
                        step.Expected = StepMatrix[row - 2, col - 1];
                    }

                    else if (col % 4 == 0)
                    {
                        step.Program = StepMatrix[row - 2, col - 1];
                    }

                }
                //Add the Step object to the List of object to get a full table
                stepList.Add(step);

            }

            return stepList;
        }

        
        //***********************************************************************//Extract Test Description Method//*******************************************************************//
        public TestDescriptionTable ExtractTestDescription(string path)
        {
            short tableIndex = 1;

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);
            
            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);
            
            //Instanciate a list of Test Description object  
            TestDescriptionTable testDescription = new TestDescriptionTable();

            //Define the table that we will be working on
            Table Table1 = allDocument.Tables[tableIndex];

            //Instanciate Test Description object to fill it every time we go over a column
            TestDescriptionTable testDescriptionTable1 = new TestDescriptionTable();

            //Looping over the rows while fixing the columns 
            for (int row = 1; row <= Table1.Rows.Count; row++)
                
                //Switching from case to another everytime rows value changes
                switch (row)
                {
                    case 1:
                        //Setting a specific cell to a variable
                        var documentAuthorCell = Table1.Cell(row, 2);
                        //Filling a variable with the cell content
                        testDescription.Author = ReplaceNonPrintableCharacters(documentAuthorCell.Range.Text);
                        break;

                    case 2:
                        var testCampaignCell = Table1.Cell(row, 2);
                        if (ReplaceNonPrintableCharacters(testCampaignCell.Range.Text) == "Input targeted campaign")
                            testDescription.TestCampaign = "";
                        else 
                        testDescription.TestCampaign = ReplaceNonPrintableCharacters(testCampaignCell.Range.Text);
                        break;

                    case 3:
                        var apsysDomainCell = Table1.Cell(row, 2);
                        testDescription.ApsysDomain = ReplaceNonPrintableCharacters(apsysDomainCell.Range.Text);
                        break;
                    case 4:

                        var ApplicationAmbitCell = Table1.Cell(row, 2);
                        testDescription.ApplicationAmbitApsys = ReplaceNonPrintableCharacters(ApplicationAmbitCell.Range.Text);
                        break;

                    case 5:
                        var creationCell = Table1.Cell(row, 2);
                        testDescription.Creation = ReplaceNonPrintableCharacters(creationCell.Range.Text);
                        break;

                    case 6:
                        var distributionListCell = Table1.Cell(row, 2);
                        testDescription.DistributionList = ReplaceNonPrintableCharacters(distributionListCell.Range.Text);
                        break;

                    case 7:
                        var DocumenttatusCell = Table1.Cell(row, 2);
                        testDescription.DocumentStatus = ReplaceNonPrintableCharacters(DocumenttatusCell.Range.Text);
                        break;
                }

            return testDescription;
        }

        
        //***********************************************************************//Extract Document Version Method//*******************************************************************//
        public DocumentVersionTable ExtractDocumentVersion(string path)
        {

            short tableIndex = 2;

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);
            
            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Instanciate a list of Document Version object  
            DocumentVersionTable documentVersion = new DocumentVersionTable();

            //Define the table that we will be working on
            Table Table2 = allDocument.Tables[tableIndex];

            //Looping over the columns while fixing the rows 
            for (int col = 1; col <= Table2.Columns.Count; col++)

                //Switching from case to another everytime column value changes
                switch (col)
                {
                    case 1:
                        //Setting a specific cell to a variable
                        var documentVersionCell = Table2.Cell(2, col);
                        //Filling a variable with the cell content
                        if (ReplaceNonPrintableCharacters(documentVersionCell.Range.Text) == "Click or tap here to enter text.")
                        {
                            documentVersion.Version = "";
                        }
                        else
                        {
                            documentVersion.Version = ReplaceNonPrintableCharacters(documentVersionCell.Range.Text);
                        }
                        break;
                    case 2:
                        var documentVersionDateCell = Table2.Cell(2, col);
                        documentVersion.Date = ReplaceNonPrintableCharacters(documentVersionDateCell.Range.Text);
                        break;
                    case 3:
                        var documentVersionAuthorCell = Table2.Cell(2, col);
                        if (ReplaceNonPrintableCharacters(documentVersionAuthorCell.Range.Text) == "Trigram list")
                        {
                            documentVersion.Author = "";
                        }
                        else
                        {
                            documentVersion.Author = ReplaceNonPrintableCharacters(documentVersionAuthorCell.Range.Text);
                        }
                        break;
                    case 4:
                        var documentVersionNotesCell = Table2.Cell(2, col);
                        documentVersion.Notes = ReplaceNonPrintableCharacters(documentVersionNotesCell.Range.Text);
                        break;
                }

            return documentVersion;
        }

        
        //***********************************************************************//Extract References Table Method//*******************************************************************//
        public ReferencesTable ExtractReferencesTable(string path)
        {

            short tableIndex = 3;

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);
            
            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Instanciate a list of referencesTable object  
            ReferencesTable referencestable = new ReferencesTable();

            //Define the table that we will be working on
            Table Table3 = allDocument.Tables[tableIndex];

            //Looping over the columns while fixing the rows 
            for (int col = 1; col <= Table3.Columns.Count; col++)

                //Switching from case to another everytime column value changes
                switch (col)
                {
                    case 1:
                        //Setting a specific cell to a variable
                        var referencesDocumentCell = Table3.Cell(2, col);
                        //Filling a variable with the cell content
                        referencestable.Document = ReplaceNonPrintableCharacters(referencesDocumentCell.Range.Text);
                        break;
                    case 2:
                        var referencesContentCell = Table3.Cell(2, col);
                        referencestable.Content = ReplaceNonPrintableCharacters(referencesContentCell.Range.Text);
                        break;


                }
            return referencestable;

        }

        
        //***********************************************************************//Extract Validation Table Method//*******************************************************************//
        public ValidationTable ExtractValidationTable(string path)
        {
            short tableIndex = 4;

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);

            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);
            
            //Instanciate a list of Validation Table object  
            ValidationTable validationTable = new ValidationTable();

            //Define the table that we will be working on
            Table Table4 = allDocument.Tables[tableIndex];

            //Looping over the columns while fixing the rows 
            for (int col = 0; col <= Table4.Columns.Count; col++)
                //Switching from case to another everytime column value changes
                switch (col)
                {
                    case 1:
                        //Setting a specific cell to a variable
                        var validationDateCell = Table4.Cell(2, col);

                        //Filling a variable with the cell content
                        validationTable.Date = ReplaceNonPrintableCharacters(validationDateCell.Range.Text);
                        break;

                    case 2:
                        var validationWhoCell = Table4.Cell(2, col);
                        validationTable.Who = ReplaceNonPrintableCharacters(validationWhoCell.Range.Text);
                        break;

                    case 3:
                        var validationNoteCell = Table4.Cell(2, col);
                        validationTable.Note = ReplaceNonPrintableCharacters(validationNoteCell.Range.Text);
                        break;

                }

            return validationTable;

        }

        
        //**************************************************************************//Extract Folder Table Method//*******************************************************************//
        public FolderTable ExtractFolderValue(string path)
        {
            short tableIndex = 5;

            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);

            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Instanciate a list of Folder Table object  
            FolderTable folderTable = new FolderTable();

            //Define the table that we will be working on
            Table Table5 = allDocument.Tables[tableIndex];

            //Looping over the columns while fixing the rows 
            for (int col = 0; col <= Table5.Columns.Count; col++)
                
                //Switching from case to another everytime column value changes
                switch (col)
                {
                    case 1:
                        //Setting a specific cell to a variable
                        var folderNameCell = Table5.Cell(2, col);

                        //Filling a variable with the cell content
                        folderTable.Parent = ReplaceNonPrintableCharacters(folderNameCell.Range.Text);
                        break;
                    case 2:
                        var folderValueCell = Table5.Cell(2, col);
                        folderTable.Folder = ReplaceNonPrintableCharacters(folderValueCell.Range.Text);
                        break;
                }

            return folderTable;
        }

        
        //****************************************************************************//Extract Test Cases Method//*******************************************************************//
        public List<TestCase> ExtractTestCases(string path)
        {
            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);
            
            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Defining a range for all test cases
            int start = wordDocument.Content.Text.IndexOf("If possible give the list of useful parameter and their combination") + "If possible give the list of useful parameter and their combination".Length;
            int end = wordDocument.Content.Text.LastIndexOf("***") + "***".Length;

            //Select Test Cases Range by providing to start index and the length of the range
            string testcaseRange = wordDocument.Content.Text.Substring(start, end - start);

            //Create an instance of the Test Case List
            List<TestCase> TestCaseList = new List<TestCase>();

            //Counting how much test cases are in the document
            var wordToFind = "Test_case";
            var testcaseCounter = 0;

            //Loop over every Test Case in the document
            foreach (Match match in Regex.Matches(allDocument.Text, wordToFind))
            {
                //Incremente Test Case Counter for every match of the Test Case word with the file words
                testcaseCounter++;
            }

            //*************************************************************************************************
            //// creating Excel Application
            //Excel.Application xlApp = new Excel.Application();
            //// creating new WorkBook within Excel application
            //Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(Type.Missing);
            //// creating new Excelsheet in workbook
            //Excel.Worksheet xlWorkSheet = new Excel.Worksheet();


            ////The path where the app will save the excel file
            //string excelPath = path.Replace("dotx", "csv");

            //Excel.Sheets xlSheets = null;

            //xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            //// see the excel sheet behind the program
            //xlApp.Visible = false;
            //*************************************************************************************************

            //Set the Step table index as it's loctation on the word file
            int indiceStep = 14;
            int indicePreCondition = 13;
            
            //Loop over test case and extract all it's information
            for (int i = 1; i < testcaseCounter + 1; i++)

            {      
                // I is the index of the current test case and J is the index of the next test case
                int j = i + 1;

                //*************************************************************************************************
                ////xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);

                //xlWorkSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

                //xlWorkSheet.Name = "TestCase_" + i;

                //xlWorkSheet.SaveAs(excelPath);
                //*************************************************************************************************

                //If statement that selects the last test case
                if (i == testcaseCounter)
                {

                    //Create and instance of the test case object
                    TestCase testCase = new TestCase();

                    //Set the starting index of the "i" test case
                    int startCase = wordDocument.Content.Text.IndexOf("Test_case " + i + ":");

                    //Set the end of the range as the end of the document
                    int endCase = end;

                    //Define a range for this test case 
                    string singleCaseRange = wordDocument.Content.Text.Substring(startCase, endCase - startCase);

                    //Extracting Title
                    testCase.Title = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Test_case " + i + ":", "Characteristics"));

                    //Extracting Context 
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "of circumstances.", "Aim:")) == "Input text.")
                    {
                        testCase.Context = "";
                    }
                    else
                    {
                        testCase.Context = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "of circumstances.", "Aim:"));

                    }


                    //Extracting Aim
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Define the goal.", "Pre-condition:")) == "Input text")
                    {
                        testCase.Aim = "";
                    }
                    else
                    {
                        testCase.Aim = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Define the goal.", "Pre-condition:"));
                    }

                    //Extracting Post-condition
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "after the execution.", "Step:")) == "Input Items")
                    {
                        testCase.PostCondition = "";
                    }
                    else
                    {
                        testCase.PostCondition = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "after the execution.", "Step:"));
                    }
                    testCase.PostCondition = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "after the execution.", "Step:"));

                    //Extracting Repetition
                    testCase.Repetition = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "iteration number.", "***"));

                    //Extracting Test case Characteristics
                    testCase.TestCaseCaracteristics = new TestCaseCaracteristics();

                    //If the user don't choose a value for this field return an empty string
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Type:", "Id")) == " Choose an item.")
                    {
                        testCase.TestCaseCaracteristics.TypeQC = "";
                    }

                    else
                    {
                        testCase.TestCaseCaracteristics.TypeQC = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Type:", "Id"));

                    }

                    //If the user don't choose a value for this field return an empty string
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Priority:", "Context:")) == " Choose an item.")
                    {
                        testCase.TestCaseCaracteristics.Priority = "";
                    }
                    else
                    {
                        testCase.TestCaseCaracteristics.Priority = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Priority:", "Context:"));

                    }

                    int val;

                    //Extract test case ID
                    if (int.TryParse(ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Id:", "Priority:")), out val) == true)
                        testCase.TestCaseCaracteristics.Id = Convert.ToInt32(ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Id:", "Priority:")));
                    else
                        testCase.TestCaseCaracteristics.Id = 0;

                    //Extract PreCondition Table from test case using the Extract PreCondition Method
                    testCase.PreCondition = ExtractPrecondition(path, indicePreCondition);
                    indicePreCondition += 3;
                    //Extract Step Table from test case using the Extract Step Method
                    testCase.Step = ExtractStep(path, indiceStep);
                    indiceStep += 3;

                    //testCase.Configuration = ExtractConfiguration(path, xlWorkSheet);
                    MapTestCaseSteps.Add(i, testCase.Step.Count);

                    //Add this extract test case to the list of test cases objects
                    TestCaseList.Add(testCase);

                }

                //Else is for extracting every test case except the last one
                else

                {
                    //Create and instance of the test case object
                    TestCase testCase = new TestCase();

                    //Set the starting index of the "i" test case
                    int startCase = wordDocument.Content.Text.IndexOf("Test_case " + i + ":");

                    //Set the starting index of the "j" test case [Next Test Case]
                    int endCase = wordDocument.Content.Text.IndexOf("Test_case " + j + ":");

                    //Define a range for this test case 
                    string singleCaseRange = wordDocument.Content.Text.Substring(startCase, endCase - startCase);

                    //Extracting Title
                    testCase.Title = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Test_case " + i + ":", "Characteristics"));

                    //Extracting Context 
                    testCase.Context = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "of circumstances.", "Aim:"));

                    //Extracting Aim
                    testCase.Aim = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Define the goal.", "Pre-condition:"));

                    //Extracting Post-condition
                    testCase.PostCondition = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "after the execution.", "Step:"));

                    //Extracting Repetition
                    testCase.Repetition = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "iteration number.", "***"));

                    //Extracting Test case Characteristics
                    testCase.TestCaseCaracteristics = new TestCaseCaracteristics();

                    //If the user don't choose a value for this field return an empty string
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Type: ", "Id")) == " Choose an item.")
                    {
                        testCase.TestCaseCaracteristics.TypeQC = "";
                    }
                    else
                    {
                        testCase.TestCaseCaracteristics.TypeQC = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Type: ", "Id"));
                        testCase.TestCaseCaracteristics.TypeQC = testCase.TestCaseCaracteristics.TypeQC.TrimStart();

                    }

                    //If the user don't choose a value for this field return an empty string
                    if (ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Priority:", "Context:")) == " Choose an item.")
                    {
                        testCase.TestCaseCaracteristics.Priority = "";
                    }
                    else
                    {
                        testCase.TestCaseCaracteristics.Priority = ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Priority:", "Context:"));

                    }

                    int val;

                    //Extract test case Id
                    if (int.TryParse(ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Id:", "Priority:")), out val) == true)
                        testCase.TestCaseCaracteristics.Id = Convert.ToInt32(ReplaceNonPrintableCharacters(GetBetween(singleCaseRange, "Id:", "Priority:")));
                    else
                        testCase.TestCaseCaracteristics.Id = 0;

                    //Extract PreCondition Table from test case using the Extract PreCondition Method
                    testCase.PreCondition = ExtractPrecondition(path, indicePreCondition);
                    indicePreCondition += 3;

                    //Extract Step Table from test case using the Extract Step Method
                    testCase.Step = ExtractStep(path, indiceStep);
                    indiceStep += 3;

                    //testCase.Configuration = ExtractConfiguration(path, xlWorkSheet);

                    //Add this extract test case to the list of test cases objects
                    MapTestCaseSteps.Add(i, testCase.Step.Count);

                    TestCaseList.Add(testCase);
                }



            }



            //*************************************************************************************************
            ////Save the excel file 
            //xlWorkBook.SaveAs(excelPath, Excel.XlFileFormat.xlCSVMSDOS, Excel.XlSaveAsAccessMode.xlNoChange);

            ////Close the Workbook and quit the excel Application
            //xlWorkBook.Close(SaveChanges: true, Filename: excelPath);
            //xlApp.Quit();
            //*************************************************************************************************


            return TestCaseList;
        }


        //****************************************************************************//Extract Scenario Method//*********************************************************************//
        public Scenario ExtractScenario(string path)
        {
            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);
            
            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Create an instance of Scenario method
            Scenario scenario = new Scenario();

            //Defining a range for scenario section
            string scenarioRange = GetBetween(wordDocument.Content.Text, "Scenario", "Test case Grid");

            //Extracting Scenario Characteristics
            scenario.Context = ReplaceNonPrintableCharacters(GetBetween(scenarioRange, "Context:", "Type:"));

            //Extract Type
            scenario.Type = ReplaceNonPrintableCharacters(GetBetween(scenarioRange, "Type:", "Target:"));

            //Extract Target 
            scenario.Target = ReplaceNonPrintableCharacters(GetBetween(scenarioRange, "Target:", "Research Orientation:"));

            //Extract Research Orientation
            scenario.ResearchOrientation = ReplaceNonPrintableCharacters(GetBetween(scenarioRange, "Research Orientation:", "Environment:"));

            //Extract environment
            scenario.Environment = ReplaceNonPrintableCharacters(GetBetween(scenarioRange, "Environment:", "Test case Grid"));

            //Extract Test Cases
            scenario.TestCases = ExtractTestCases(path);

            return scenario;
        }

        //******************************************************************************//Extract Global File Method//**************************************************************//
        public FunctionalTestDocument ExtractGlobalFile(string path)
        {
            //Open the word document on a ReadOnly mode without displaying it
            Document wordDocument = Application.Documents.Open(path, ReadOnly: true, Visible: false);

            log.Info("*******************************************************************************************************************");

            log.Info("The application started");

            //Set a range for all the word file to manage it
            Range allDocument = wordDocument.Range(0, wordDocument.Content.Characters.Count);

            //Create an instance of globalobject Object 
            FunctionalTestDocument globalObject = new FunctionalTestDocument();

            // Getting the document ID from the word document
            if (ReplaceNonPrintableCharacters(wordDocument.Sentences[2].Text) == "Project ID")
                {
                globalObject.DocumentId = "";
                }
            else
            {
                globalObject.DocumentId = ReplaceNonPrintableCharacters(wordDocument.Sentences[2].Text);

            }
            // Getting the document ID from the word document

            globalObject.DocumentText = ReplaceNonPrintableCharacters(wordDocument.Sentences[3].Text);

            //Extract Test Description table
            globalObject.TestDescriptionTable = ExtractTestDescription(path);
            log.Info("Extract Test Description table.");
            //Extract Document version table
            globalObject.DocumentVersionTable = ExtractDocumentVersion(path);
            log.Info("Extract Test Description table.");
            //Extract Reference table
            globalObject.ReferencesTable = ExtractReferencesTable(path);
            log.Info("Extract Reference table.");
            //Extract Validation Table
            globalObject.ValidationTable = ExtractValidationTable(path);
            log.Info("Extract Validation Table.");
            //Extract Folder Table
            globalObject.FolderTable = ExtractFolderValue(path);
            log.Info("Extract Folder Table.");
            //Extract Scenario
            globalObject.Scenario = ExtractScenario(path);
            log.Info("Extract Scenario.");

            //Close the word document and quit the word application too
            wordDocument.Close();
            log.Info("Close Word Document.");
            //Quit the application
            Application.Quit();
            log.Info("Quit Application.");
            GC.Collect();
            // Waiting till finilizer thread will call all finalizers
            GC.SuppressFinalize(Application);


            return globalObject;
        }

        //***********************************************************************//Replace non printable Characters Method//********************************************************//
        public static string ReplaceNonPrintableCharacters(string value)
        {
            //Define the range of pattern to remove
            string pattern = "[^ -~]+";

            //create an Instance of these patterns
            Regex reg_exp = new Regex(pattern);

            //replace everytime there is a pattern by a void string
            return reg_exp.Replace(value, "");
        }

        
        //********************************************************************************//GetBetween Method//*********************************************************************//
        public string GetBetween(string strSource, string strStart, string strEnd)
        {
            //Start and End are key parameters to delimit the range to choose 
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }


        //****************************************************************************//Get Test Cases Method//*********************************************************************//
        public List<JsonModelEmpty> GetTestCases(FunctionalTestDocument globalObject, List<int> typeQCList)
        {
            //Create instance of Json Model List 
            List<JsonModelEmpty> JsonModelList = new List<JsonModelEmpty>();

            //Get the number of test cases from a matrix that holds test cases as key and number of steps as a value
            int caseCounter = MapTestCaseSteps.Keys.Count;

            //Get the document author name from the word file
            string userName = globalObject.TestDescriptionTable?.Author;

            //Get the Json users list using a get methode
            string usersJson = GetQACUsers();
            
            //Search for user Id by going over the list of users
            int authorId = SearchUserId(usersJson, userName);

            //Get folder name from the word file
            string folderName = globalObject.FolderTable?.Folder;

            string parentName = globalObject.FolderTable?.Parent;
            
            //Get all folders list
            string foldersJson = GetQACFolders(11873);

            //Search for folder Id by going over the list of folders
            int dossierId = SearchFolderId(foldersJson, folderName, parentName);

            //Loop over all test cases
            for (int iCase = 0; iCase < caseCounter; iCase++)
            {
                int stepCounter = 0;
                // Get the number of steps from a matrix
                stepCounter = MapTestCaseSteps[iCase + 1];

                //MapTestCaseSteps.TryGetValue(iCase, out stepCounter);
                List<JsonStep> JsonStepList = new List<JsonStep>();
                List<CustomFields> CustomList = new List<CustomFields>();

                //Loop over the steps
                for (var iStep = 0; iStep < stepCounter; iStep++)
                {
                    JsonStep stp = new JsonStep()
                    {
                        ExpectedResult = globalObject.Scenario.TestCases[iCase]?.Step[iStep]?.Expected,
                        Seq = globalObject.Scenario.TestCases[iCase]?.Step[iStep]?.StepNum,
                        Step = globalObject.Scenario.TestCases[iCase]?.Step[iStep]?.Description,
                    };
                    if (!stp.IsEmpty())
                    {
                        JsonStepList.Add(stp);
                    }

                }

                PutCustomFields(globalObject, iCase, stepCounter, CustomList);
                string Title = globalObject.Scenario?.TestCases[iCase]?.Title;
                string Description = "<b>Context:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.Context
                          + " <br />" + " <br />" + "<b>Aim:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.Aim
                          + " <br />" + " <br />" + "<b>PreCondition:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.PreCondition
                          + " <br />" + " <br />" + "<b>PostCondition:</b> " + " <br />" + globalObject.Scenario?.TestCases[iCase]?.PostCondition;
                string TestType = globalObject.Scenario?.TestCases[iCase]?.TestCaseCaracteristics?.TypeQC.TrimStart();
                string Priority = globalObject.Scenario?.TestCases[iCase]?.TestCaseCaracteristics?.Priority.TrimStart();

                    typeQCList.Add(iCase);


                JsonModelList.Add(JsonModelEmpty.BuildInstance(globalObject, iCase, authorId, dossierId, Title, true, Description, "New", TestType, Priority, JsonStepList, CustomList));


            }

            return JsonModelList;

        }


        //****************************************************************************//Create Test Method//*********************************************************************//
        public void CreateTest(FunctionalTestDocument globalObject)
        {
            int projectId = 11874;
            List<JsonModelEmpty>  JsonList = new List<JsonModelEmpty>();
            List<int> typeQCList = new List<int>();           

            JsonList = GetTestCases(globalObject, typeQCList);

            //foreach(JsonModelEmpty JM in JsonList)
            for (int i = 0; i < JsonList.Count; i++)

            {
                int testId;
                testId =globalObject.Scenario.TestCases[typeQCList[i]].TestCaseCaracteristics.Id;
                JsonModelEmpty JM = JsonList[i];
                // Convert Object to Json 
                string JsonModelfile = new JavaScriptSerializer().Serialize(JM);

                if (globalObject.Scenario.TestCases[typeQCList[i]].TestCaseCaracteristics.TypeQC.Trim() == "Creation")
                    PostTest(projectId, JsonModelfile);
                else if (globalObject.Scenario.TestCases[typeQCList[i]].TestCaseCaracteristics.TypeQC.Trim() == "Update")
                    PutTest(projectId, JsonModelfile, testId);
                else
                    log.Error("The Type field should take either Creation or Update");


            }

        }

        
        //****************************************************************************//Post Test Cases Method//*********************************************************************//
        public void PostTest(int projectId, string json)
        {
            // Create a request using a URL that can receive a post. 
            string URL = "http://emeagvaqc01.newaccess.ch/rest-api/service/api/v1/projects/"+ projectId +"/tests";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
           
            // Set the Method property of the request to POST.
            request.Method = "POST";

            request.PreAuthenticate = true;
            request.Headers.Add("Authorization", "Basic YWRtaW5RQUNAbmV3YWNjZXNzLmNoOlFBX05XQTIwMTg=");
            request.ContentType = "application/json";

            request.ContentLength = Encoding.UTF8.GetByteCount(json);
            StreamWriter streamWriter = null;
            streamWriter = new StreamWriter(request.GetRequestStream());
            streamWriter.Write(json);
            streamWriter.Close();

            WebResponse response = request.GetResponse();
            var httpResponse = (HttpWebResponse)request.GetResponse();
            var streamReader = new StreamReader(httpResponse.GetResponseStream());
            var result = streamReader.ReadToEnd();
            // System.Console.WriteLine(result);
            streamReader.Close();
            httpResponse.Close();

        }

        //****************************************************************************//Put Test Cases Method//*********************************************************************//
        public void PutTest(int projectId, string json, int testId)
        {
            // Create a request using a URL that can receive a put. 
            string URL = "http://emeagvaqc01.newaccess.ch/rest-api/service/api/v1/projects/" + projectId + "/tests/" + testId;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);

            // Set the Method property of the request to PUT.
            request.Method = "PUT";

            request.PreAuthenticate = true;
            request.Headers.Add("Authorization", "Basic YWRtaW5RQUNAbmV3YWNjZXNzLmNoOlFBX05XQTIwMTg=");
            request.ContentType = "application/json";

            request.ContentLength = Encoding.UTF8.GetByteCount(json);
            StreamWriter streamWriter = null;
            streamWriter = new StreamWriter(request.GetRequestStream());
            streamWriter.Write(json);
            streamWriter.Close();

            WebResponse response = request.GetResponse();
            var httpResponse = (HttpWebResponse)request.GetResponse();
            var streamReader = new StreamReader(httpResponse.GetResponseStream());
            var result = streamReader.ReadToEnd();
            // System.Console.WriteLine(result);
            streamReader.Close();
            httpResponse.Close();

        }

        //****************************************************************************//GetQAC Users Cases Method//*********************************************************************//
        public string GetQACUsers()
        {
            int deptId = 8162;

            string url = "http://emeagvaqc01.newaccess.ch/rest-api/service/api/v1/depts/" + deptId + "/users";

            var request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = "GET";

            request.PreAuthenticate = true;
            request.Headers.Add("Authorization", "Basic YWRtaW5RQUNAbmV3YWNjZXNzLmNoOlFBX05XQTIwMTg=");
            request.ContentType = "application/json";

            var content = string.Empty;

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                using (var stream = response.GetResponseStream())
                {
                    using (var sr = new StreamReader(stream))
                    {
                        content = sr.ReadToEnd();
                    }
                }
            }

            return content;

        }

        //****************************************************************************//SerchUserId Cases Method//*********************************************************************//
        public int SearchUserId(string json, string name)
            
        {
            string email = name + "@newaccess.ch";

            JObject data = JObject.Parse(json);

            JArray results = (JArray)data["results"];

            var result =
                from user in results.Where(i => (string)i["email"] == email)

                select new
                {
                    id = (int)user["id"]
                };
            

            return result.FirstOrDefault().id;


        }

        //****************************************************************************//GetQAC folders Cases Method//*********************************************************************//
        public string GetQACFolders(int projectId)

        {
            string url = "http://emeagvaqc01.newaccess.ch/rest-api/service/api/v2/projects/" + projectId + "/Tests/folders";

            var request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = "GET";

            request.PreAuthenticate = true;

            request.Headers.Add("Authorization", "Basic YWRtaW5RQUNAbmV3YWNjZXNzLmNoOlFBX05XQTIwMTg=");

            request.ContentType = "application/json";

            var content = string.Empty;

            using (var response = (HttpWebResponse)request.GetResponse())

            {
                using (var stream = response.GetResponseStream())

                {
                    using (var sr = new StreamReader(stream))
                    {
                        content = sr.ReadToEnd();
                    }
                }
            }
            return content;
        }

        //****************************************************************************//Search Folder Id Cases Method//*********************************************************************//
        public int SearchFolderId(string json, string foldername , string parentname)
        {
            JObject data = JObject.Parse(json);
            JArray results = (JArray)data["results"];

            var result =
                from folder in results.Where(i => ((string)i["folder_name"] == foldername) && ((string)i["parent_name"] == parentname) )
                select new
                {
                    id = (int)folder["id"],
                    parent_id = (int)folder["parent_id"]
                };

            return result.FirstOrDefault().id;
        }


        public void PutCustomFields(FunctionalTestDocument globalObject, int iCase, int stepCounter, List<CustomFields> CustomFieldslist)

        {
            string programs = "";

            List<CustomFields> list = new List<CustomFields>();
            list.Add(CustomFields.SetCustomFields("Custom3", "Project id", globalObject.DocumentId));
            list.Add(CustomFields.SetCustomFields("Custom2", "Release", globalObject.TestDescriptionTable?.TestCampaign));
            for (var iStep = 0; iStep < stepCounter; iStep++)
            {
                programs += globalObject.Scenario?.TestCases[iCase]?.Step[iStep]?.Program + (iStep==stepCounter-1?"": ",");
            };

            list.Add(CustomFields.SetCustomFields("Custom4", "Program", programs));
            list.Add(CustomFields.SetCustomFields("Custom5", "Author", globalObject.DocumentVersionTable?.Author));
            list.Add(CustomFields.SetCustomFields("Custom1", "Keyword", globalObject.Scenario?.ResearchOrientation));


            CustomFieldslist.AddRange(list);


        }

        public string GetName()
        {
            throw new NotImplementedException();
        }

        public void UploadFile(string path)
        {
            throw new NotImplementedException();
        }

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
    }
}
