namespace test
{
    using Microsoft.Office.Interop.Word;
    using System;
    using System.Reflection;
    using Excel = Microsoft.Office.Interop.Excel;

    class Program
    {
        static void Main(string[] args)
        {
            //Instanciate application as the app we will be working <ith during the project
            Application application = new Application();
            
            //Path of Word document
            object fileName = @"C:\Users\sadok.jbenyeni\Downloads\TemplateTest.docx";
            
            //Declaration of an object for the unknowing status of some parameters on open method
            object nullobj = Missing.Value;

            //Opening the word file
            Document document = application.Documents.Open(ref fileName, ref nullobj, ReadOnly: false, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);

            //word visibility set to false so the word file didn't pop on he screen
            application.Visible = false;
            
            // All document is the range of all the document content
            Range allDocument = document.Range(0, document.Content.Characters.Count);

            // Getting the document ID from the word document
            string documentID = document.Sentences[2].Text;
            //This method removes the non printable characters
            documentID = ReplaceCharacters.ReplaceNonPrintableCharacters(documentID);

            // Getting the document ID from the word document
            string documentText = document.Sentences[3].Text;
            documentText = ReplaceCharacters.ReplaceNonPrintableCharacters(documentText);

            //------------------------------------------Test Description Table--------------------------------------------------------------

            //Declaring a table by his position in the document
            Table Table1 = allDocument.Tables[1];
            TestDescriptionTable testDescriptionTable1 = new TestDescriptionTable();
            //Looping through the rows while fixing the columns 
            for (int row = 1; row <= Table1.Rows.Count; row++)
                //Switching from case to another everytime rows value changes
                switch (row)
                {
                    case 1:
                        //Setting a specific cell to a variable
                        var documentAuthorCell = Table1.Cell(row, 2);
                        //Filling a variable with the cell content
                        var documentAuthor = documentAuthorCell.Range.Text;
                        documentAuthor = ReplaceCharacters.ReplaceNonPrintableCharacters(documentAuthor);
                        break;
                    case 2:
                        var testCampaignCell = Table1.Cell(row, 2);
                        var testCampaign = testCampaignCell.Range.Text;
                        testCampaign = ReplaceCharacters.ReplaceNonPrintableCharacters(testCampaign);
                        break;
                    case 3:
                        var apsysDomainCell = Table1.Cell(row, 2);
                        var apsysDomain = apsysDomainCell.Range.Text;
                        apsysDomain = ReplaceCharacters.ReplaceNonPrintableCharacters(apsysDomain);
                        break;
                    case 4:
                        var applicationAmbitCell = Table1.Cell(row, 2);
                        var applicationAmbit = applicationAmbitCell.Range.Text;
                        applicationAmbit = ReplaceCharacters.ReplaceNonPrintableCharacters(applicationAmbit);
                        break;
                    case 5:
                        var creationCell = Table1.Cell(row, 2);
                        var creation = creationCell.Range.Text;
                        creation = ReplaceCharacters.ReplaceNonPrintableCharacters(creation);
                        break;
                    case 6:
                        var distributionListCell = Table1.Cell(row, 2);
                        var distributionList = distributionListCell.Range.Text;
                        distributionList = ReplaceCharacters.ReplaceNonPrintableCharacters(distributionList);
                        break;
                    case 7:
                        var documentStatusCell = Table1.Cell(row, 2);
                        var documentStatus = documentStatusCell.Range.Text;
                        documentStatus = ReplaceCharacters.ReplaceNonPrintableCharacters(documentStatus);
                        break;
                }

            //------------------------------------------Document Version Table--------------------------------------------------------------


            Table Table2 = allDocument.Tables[2];
            for (int col = 1; col <= Table2.Columns.Count; col++)

                switch (col)
                {
                    case 1:
                        var documentVersionCell = Table2.Cell(2, col);
                        var documentVersion = documentVersionCell.Range.Text;
                        documentVersion = ReplaceCharacters.ReplaceNonPrintableCharacters(documentVersion);
                        break;
                    case 2:
                        var documentVersionDateCell = Table2.Cell(2, col);
                        var documentVersionDate = documentVersionDateCell.Range.Text;
                        documentVersionDate = ReplaceCharacters.ReplaceNonPrintableCharacters(documentVersionDate);
                        break;
                    case 3:
                        var documentVersionAuthorCell = Table2.Cell(2, col);
                        var documentVersionAuthor = documentVersionAuthorCell.Range.Text;
                        documentVersionAuthor = ReplaceCharacters.ReplaceNonPrintableCharacters(documentVersionAuthor);
                        break;
                    case 4:
                        var documentVersionNotesCell = Table2.Cell(2, col);
                        var documentVersionNotes = documentVersionNotesCell.Range.Text;
                        documentVersionNotes = ReplaceCharacters.ReplaceNonPrintableCharacters(documentVersionNotes);
                        break;
                }

            //-----------------------------------------References Table ---------------------------------------------------------


            Table Table3 = allDocument.Tables[3];

            for (int col = 1; col <= Table3.Columns.Count; col++)
                switch (col)
                {
                    case 1:
                        var referencesDocumentCell = Table3.Cell(2, col);
                        var referencesDocument = referencesDocumentCell.Range.Text;
                        referencesDocument = ReplaceCharacters.ReplaceNonPrintableCharacters(referencesDocument);
                        break;
                    case 2:
                        var referencesContentCell = Table3.Cell(2, col);
                        var referencesContent = referencesContentCell.Range.Text;
                        referencesContent = ReplaceCharacters.ReplaceNonPrintableCharacters(referencesContent);
                        break;


                }

            //------------------------------------------ Validation Table ------------------------------------------------------------
            Table Table4 = allDocument.Tables[4];
            for (int col = 0; col <= Table4.Columns.Count; col++)
                switch (col)
                {
                    case 1:
                        var validationDateCell = Table4.Cell(2, col);
                        var validationDate = validationDateCell.Range.Text;
                        validationDate = ReplaceCharacters.ReplaceNonPrintableCharacters(validationDate);
                        break;
                    case 2:
                        var validationWhoCell = Table4.Cell(2, col);
                        var validationWho = validationWhoCell.Range.Text;
                        validationWho = ReplaceCharacters.ReplaceNonPrintableCharacters(validationWho);
                        break;
                    case 3:
                        var validationNoteCell = Table4.Cell(2, col);
                        var validationNote = validationNoteCell.Range.Text;
                        validationNote = ReplaceCharacters.ReplaceNonPrintableCharacters(validationNote);
                        break;

                }


            //Defining a range for scenario section
            string scenarioRange = Between.GetBetween(document.Content.Text, "Scenario - ", "If possible give the list of useful parameter and their combination");

            //Extracting Scenario Characteristics
            string scenarioContext = Between.GetBetween(scenarioRange, "Context:", "Type:");
            string scenarioType = Between.GetBetween(scenarioRange, "Type:", "Target:");
            string scenarioTarget = Between.GetBetween(scenarioRange, "Target:", "Research Orientation:");
            string scenarioResearchOrientation = Between.GetBetween(scenarioRange, "Research Orientation:", "Environment:");
            string scenarioEnvironment = Between.GetBetween(scenarioRange, "Environment:", "Test case Grid");
            scenarioContext = ReplaceCharacters.ReplaceNonPrintableCharacters(scenarioContext);
            scenarioType = ReplaceCharacters.ReplaceNonPrintableCharacters(scenarioType);
            scenarioTarget = ReplaceCharacters.ReplaceNonPrintableCharacters(scenarioTarget);
            scenarioResearchOrientation = ReplaceCharacters.ReplaceNonPrintableCharacters(scenarioResearchOrientation);
            scenarioEnvironment = ReplaceCharacters.ReplaceNonPrintableCharacters(scenarioEnvironment);

            //Defining a range for test case
            int Start = document.Content.Text.IndexOf("If possible give the list of useful parameter and their combination") + "If possible give the list of useful parameter and their combination".Length;
            int End = document.Content.Text.LastIndexOf("***") + "***".Length;
            string testcaseRange = document.Content.Text.Substring(Start, End - Start);

            //Extracting Test case Characteristics
            string testCaseType = Between.GetBetween(testcaseRange, "Type QC:", "Id");
            string testCaseId = Between.GetBetween(testcaseRange, "Id:", "Level:");
            string testCaseLevel = Between.GetBetween(testcaseRange, "Level:", "Priority:");
            string testCasePriority = Between.GetBetween(testcaseRange, "Priority:", "Context:");
            testCaseType = ReplaceCharacters.ReplaceNonPrintableCharacters(testCaseType);
            testCaseId = ReplaceCharacters.ReplaceNonPrintableCharacters(testCaseId);
            testCaseLevel = ReplaceCharacters.ReplaceNonPrintableCharacters(testCaseLevel);
            testCasePriority = ReplaceCharacters.ReplaceNonPrintableCharacters(testCasePriority);

            //Extracting Context 
            string contextText = Between.GetBetween(testcaseRange, "of circumstances.", "Aim:");
            contextText = ReplaceCharacters.ReplaceNonPrintableCharacters(contextText);

            //Extracting Aim
            string aimText = Between.GetBetween(testcaseRange, "Define the goal.", "Pre-condition:");
            aimText = ReplaceCharacters.ReplaceNonPrintableCharacters(aimText);

            //Extracting Post-condition
            string postConditionText = Between.GetBetween(testcaseRange, "after the execution.", "Step:");
            postConditionText = ReplaceCharacters.ReplaceNonPrintableCharacters(postConditionText);

            //Extracting Repetition
            string repetitionText = Between.GetBetween(testcaseRange, "iteration number.", "***");
            repetitionText = ReplaceCharacters.ReplaceNonPrintableCharacters(repetitionText);



            //------------------------------------------PreCondition Table -----------------------------------------------------;

            Table Table12 = allDocument.Tables[12];
            //Declaring a matrix with the same measurment of this table
             string[,] PreConditionMatrix = new string[Table12.Rows.Count - 1, Table12.Columns.Count];

            //Going through the doc table and fill the matrix in the same order
            for (int row = 2; row <= Table12.Rows.Count; row++)
            {
                for (int col = 1; col <= Table12.Columns.Count; col++)
                {
                    var preConditionCell = Table12.Cell(row,col);
                    PreConditionMatrix[row - 2, col - 1] = preConditionCell.Range.Text;
                    PreConditionMatrix[row - 2, col - 1] = ReplaceCharacters.ReplaceNonPrintableCharacters(PreConditionMatrix[row - 2, col - 1]); 
                }

            }




            //----------------------------------------Step Table --------------------------------------------------------------;


            Table Table13 = allDocument.Tables[13];
            string[,] StepMatrix = new string[Table13.Rows.Count - 1, Table13.Columns.Count];


            for (int row = 2; row <= Table13.Rows.Count; row++)
            {
                for (int col = 1; col <= Table13.Columns.Count; col++)
                {
                    var stepCell = Table13.Cell(row, col);
                    StepMatrix[row - 2, col - 1] = stepCell.Range.Text;
                    StepMatrix[row - 2, col - 1] = ReplaceCharacters.ReplaceNonPrintableCharacters(StepMatrix[row - 2, col - 1]);
                }

            }



            //----------------------------------------Configuration Table -------------------------------------------------------------);

            // The path where the app will save the excel file
            string excelPath = @"C:\Users\sadok.jbenyeni\Downloads\TemplateTest.csv";

            //Declaring an Excel Application, Workbook and a Worksheet where the data will fit
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            //Initializing these variables
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(nullobj);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Table Table14 = allDocument.Tables[14];

            string[,] ParameterMatrix = new string[Table14.Rows.Count, Table14.Columns.Count];

            for (int col = 1; col <= Table14.Columns.Count; col++)
            {
                for (int row = 1; row <= Table14.Rows.Count; row++)
                {
                    var parameterCell = Table14.Cell(row, col);
                    ParameterMatrix[row - 1, col - 1] = parameterCell.Range.Text;
                    ParameterMatrix[row - 1, col - 1] = ReplaceCharacters.ReplaceNonPrintableCharacters(ParameterMatrix[row - 1,col - 1]);
                    //Filling the Worksheet Cells with the data from ParameterMatrix Cells
                    xlWorkSheet.Cells[row, col] = ParameterMatrix[row - 1, col - 1];

                }

            }


            //Save the excel file 
            xlWorkBook.SaveAs(excelPath, Excel.XlFileFormat.xlCSVMSDOS, nullobj, nullobj, nullobj, nullobj, Excel.XlSaveAsAccessMode.xlExclusive, nullobj, nullobj, nullobj, nullobj, nullobj);

            //Close the Workbook and quit the excel application
            xlWorkBook.Close();
            xlApp.Quit();

            //Close the word document and quit the word application too
            document.Close();
            application.Quit();


            GC.Collect();
            // Waiting till finilizer thread will call all finalizers
            GC.SuppressFinalize(application);





        }

    }
}