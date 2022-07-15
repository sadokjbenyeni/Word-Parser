using Services.Models.Tables_Folder;

namespace Services.Services
{
    public class FunctionalTestDocument
    {
        
        public int Id { get; set; }

        public string DocumentId { get; set; }

        public string DocumentText { get; set; }

        public TestDescriptionTable TestDescriptionTable { get; set; }

        public ValidationTable ValidationTable { get; set; }

        public DocumentVersionTable DocumentVersionTable { get; set; }

        public ReferencesTable ReferencesTable { get; set; }

        public FolderTable FolderTable { get; set; }

        public Scenario Scenario { get; set; }

    }
}
