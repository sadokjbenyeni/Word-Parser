namespace test
{
    public class FunctionalTestDocument
    {
        
        string DocumentId { get; set; }

        string DocumentText { get; set; }

        TestDescriptionTable TestDescriptionTable { get; set; }

        ValidationTable ValidationTable { get; set; }

        DocumentVersionTable GetDocumentVersionTable { get; set; }

        ReferencesTable ReferencesTable { get; set; }

        Scenario Scenario { get; set; }

    }
}
