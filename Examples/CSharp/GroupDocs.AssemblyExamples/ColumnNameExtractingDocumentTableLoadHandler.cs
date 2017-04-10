using GroupDocs.Assembly.Data;


namespace GroupDocs.AssemblyExamples
{
    public class ColumnNameExtractingDocumentTableLoadHandler : IDocumentTableLoadHandler
    {
        public void Handle(DocumentTableLoadArgs args)
        {
            args.Options = new DocumentTableOptions();
            args.Options.FirstRowContainsColumnNames = true;
        }
    }
}
